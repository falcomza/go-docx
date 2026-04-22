package godocx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"
)

// Updater manages a DOCX document for programmatic reading and writing.
//
// An Updater is not safe for concurrent use by multiple goroutines. All
// operations read and write files in a shared temporary directory; callers
// must serialise access externally if they need to issue operations from
// multiple goroutines.
type Updater struct {
	originalPath  string
	tempDir       string
	tempInputFile string

	bulletListNumID   int
	numberedListNumID int
	headingNumID      int

	noteRefStylesEnsured bool

	// headingNumberingDisabled suppresses <w:numPr> injection in AddHeading so that
	// the document's own heading styles (e.g. from an uploaded template) control all
	// formatting. When true, AddHeading emits only <w:pStyle> for the heading level.
	headingNumberingDisabled bool
}

// NewBlank creates a new blank DOCX document from scratch without requiring a template.
// The document contains a minimal valid OpenXML structure ready for content insertion.
func NewBlank() (*Updater, error) {
	tempDir, err := os.MkdirTemp("", "docx-blank-*")
	if err != nil {
		return nil, fmt.Errorf("create temp dir: %w", err)
	}

	if err := writeBlankDocxStructure(tempDir); err != nil {
		os.RemoveAll(tempDir)
		return nil, fmt.Errorf("write blank docx: %w", err)
	}

	u := &Updater{originalPath: "", tempDir: tempDir}

	if err := u.validateStructure(); err != nil {
		u.Cleanup()
		return nil, fmt.Errorf("invalid blank DOCX: %w", err)
	}

	return u, nil
}

// NewFromBytes creates an Updater from raw DOCX bytes (e.g., uploaded template data).
// This is useful when the template is received from a web upload, API payload, or
// database rather than a file on disk.
func NewFromBytes(data []byte) (*Updater, error) {
	if len(data) == 0 {
		return nil, errors.New("docx data is empty")
	}

	tmpFile, err := os.CreateTemp("", "docx-bytes-*.docx")
	if err != nil {
		return nil, fmt.Errorf("create temp file: %w", err)
	}
	tmpPath := tmpFile.Name()

	if _, err := tmpFile.Write(data); err != nil {
		tmpFile.Close()
		os.Remove(tmpPath)
		return nil, fmt.Errorf("write temp file: %w", err)
	}
	if err := tmpFile.Close(); err != nil {
		os.Remove(tmpPath)
		return nil, fmt.Errorf("close temp file: %w", err)
	}

	u, err := New(tmpPath)
	if err != nil {
		os.Remove(tmpPath)
		return nil, err
	}

	u.tempInputFile = tmpPath
	return u, nil
}

// New opens a DOCX file and prepares it for editing.
func New(docxPath string) (*Updater, error) {
	if docxPath == "" {
		return nil, errors.New("docx path is required")
	}
	if _, err := os.Stat(docxPath); err != nil {
		return nil, fmt.Errorf("stat docx: %w", err)
	}

	tempDir, err := os.MkdirTemp("", "docx-update-*")
	if err != nil {
		return nil, fmt.Errorf("create temp dir: %w", err)
	}

	if err := extractZip(docxPath, tempDir); err != nil {
		if rmErr := os.RemoveAll(tempDir); rmErr != nil {
			return nil, fmt.Errorf("extract docx: %w (cleanup failed: %v)", err, rmErr)
		}
		return nil, fmt.Errorf("extract docx: %w", err)
	}

	// Promote .dotx template content type to .docx document content type so that
	// callers can pass either file format transparently.
	if err := normalizeTemplateToDocument(tempDir); err != nil {
		os.RemoveAll(tempDir)
		return nil, fmt.Errorf("normalize template: %w", err)
	}

	// Normalize non-standard WordprocessingML namespace prefixes (e.g. "ns0:")
	// to the canonical "w:" prefix so all insertion logic works uniformly.
	// Documents produced by python-docx or lxml may bind the WML namespace to
	// a generated prefix such as "ns0" instead of the conventional "w".
	if err := normalizeWMLNamespacePrefix(tempDir); err != nil {
		os.RemoveAll(tempDir)
		return nil, fmt.Errorf("normalize WML namespace: %w", err)
	}

	// Normalize invalid w:characterSet values in fontTable.xml.
	// LibreOffice emits IANA charset names (e.g. "utf-8", "windows-1252") but
	// OOXML ST_UcharHexNumber requires 2-digit Windows charset IDs in hex.
	if err := normalizeFontTableCharsets(tempDir); err != nil {
		os.RemoveAll(tempDir)
		return nil, fmt.Errorf("normalize fontTable charsets: %w", err)
	}

	// Normalize invalid w:lvlJc values in numbering.xml (e.g. "start"/"end"
	// emitted by LibreOffice) so documents validate in Microsoft 365 even when
	// callers do not invoke list APIs.
	if err := normalizeNumberingLvlJc(tempDir); err != nil {
		os.RemoveAll(tempDir)
		return nil, fmt.Errorf("normalize numbering lvlJc: %w", err)
	}

	u := &Updater{originalPath: docxPath, tempDir: tempDir}

	// Validate DOCX structure
	if err := u.validateStructure(); err != nil {
		u.Cleanup()
		return nil, fmt.Errorf("invalid DOCX: %w", err)
	}

	return u, nil
}

// TempDir returns the temporary directory where the DOCX was extracted.
// It is intended for testing and advanced use only. Callers that read or write
// files directly inside TempDir bypass all validation and relationship tracking
// performed by the Updater methods.
func (u *Updater) TempDir() string {
	return u.tempDir
}

// GetChartCount returns the number of charts embedded in the document.
// Returns 0 if the document contains no charts.
func (u *Updater) GetChartCount() (int, error) {
	if u == nil {
		return 0, errors.New("updater is nil")
	}
	chartsDir := filepath.Join(u.tempDir, "word", "charts")
	entries, err := os.ReadDir(chartsDir)
	if err != nil {
		if os.IsNotExist(err) {
			return 0, nil
		}
		return 0, fmt.Errorf("read charts dir: %w", err)
	}
	var count int
	for _, e := range entries {
		if !e.IsDir() && chartFilePattern.MatchString(e.Name()) {
			count++
		}
	}
	return count, nil
}

// Cleanup removes temporary workspace.
func (u *Updater) Cleanup() error {
	if u == nil || u.tempDir == "" {
		return nil
	}
	err := os.RemoveAll(u.tempDir)
	if u.tempInputFile != "" {
		if rmErr := os.Remove(u.tempInputFile); rmErr != nil && !os.IsNotExist(rmErr) {
			if err == nil {
				err = rmErr
			}
		}
	}
	return err
}

// UpdateChart updates one chart by index (1-based).
func (u *Updater) UpdateChart(chartIndex int, data ChartData) error {
	if u == nil {
		return errors.New("updater is nil")
	}
	if chartIndex < 1 {
		return errors.New("chart index must be >= 1")
	}
	if err := validateChartData(data); err != nil {
		return err
	}

	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	if _, err := os.Stat(chartPath); err != nil {
		return fmt.Errorf("chart file does not exist: %w", err)
	}

	if err := updateChartXML(chartPath, data); err != nil {
		return fmt.Errorf("update chart xml: %w", err)
	}

	xlsxPath, err := u.findWorkbookPathForChart(chartIndex)
	if err != nil {
		return fmt.Errorf("resolve embedded workbook: %w", err)
	}
	if err := updateEmbeddedWorkbook(xlsxPath, data); err != nil {
		return fmt.Errorf("update embedded workbook: %w", err)
	}

	return nil
}

// NewFromReader opens a DOCX from an io.Reader and prepares it for editing.
// The reader content is buffered to a temporary file which is cleaned up by Cleanup().
func NewFromReader(r io.Reader) (*Updater, error) {
	if r == nil {
		return nil, errors.New("reader is nil")
	}

	tmpFile, err := os.CreateTemp("", "docx-input-*.docx")
	if err != nil {
		return nil, fmt.Errorf("create temp input file: %w", err)
	}
	tmpPath := tmpFile.Name()

	if _, err := io.Copy(tmpFile, r); err != nil {
		tmpFile.Close()
		os.Remove(tmpPath)
		return nil, fmt.Errorf("buffer reader to temp file: %w", err)
	}
	if err := tmpFile.Close(); err != nil {
		os.Remove(tmpPath)
		return nil, fmt.Errorf("close temp input file: %w", err)
	}

	u, err := New(tmpPath)
	if err != nil {
		os.Remove(tmpPath)
		return nil, err
	}

	u.tempInputFile = tmpPath
	return u, nil
}

// SaveToWriter writes the updated DOCX to an io.Writer.
func (u *Updater) SaveToWriter(w io.Writer) error {
	if u == nil {
		return errors.New("updater is nil")
	}
	if w == nil {
		return errors.New("writer is nil")
	}
	return writeZipFromDir(u.tempDir, w)
}

// Save writes the updated DOCX to outputPath.
func (u *Updater) Save(outputPath string) error {
	if u == nil {
		return errors.New("updater is nil")
	}
	if outputPath == "" {
		return errors.New("output path is required")
	}
	if err := os.MkdirAll(filepath.Dir(outputPath), 0o755); err != nil {
		return fmt.Errorf("create output dir: %w", err)
	}
	if err := createZipFromDir(u.tempDir, outputPath); err != nil {
		return fmt.Errorf("create output docx: %w", err)
	}
	return nil
}

func validateChartData(data ChartData) error {
	if len(data.Categories) == 0 {
		return errors.New("categories cannot be empty")
	}
	if len(data.Series) == 0 {
		return errors.New("series cannot be empty")
	}
	for i, s := range data.Series {
		if strings.TrimSpace(s.Name) == "" {
			return fmt.Errorf("series[%d] name cannot be empty", i)
		}
		if len(s.Values) != len(data.Categories) {
			return fmt.Errorf("series[%d] values length (%d) must match categories length (%d)", i, len(s.Values), len(data.Categories))
		}
	}
	return nil
}

func (u *Updater) findWorkbookPathForChart(chartIndex int) (string, error) {
	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	rawChart, err := os.ReadFile(chartPath)
	if err != nil {
		return "", fmt.Errorf("read chart xml for chart%d: %w", chartIndex, err)
	}

	relID := externalDataRelID(rawChart)
	if relID == "" {
		return "", fmt.Errorf("chart%d.xml has no externalData relationship ID", chartIndex)
	}

	relsPath := filepath.Join(u.tempDir, "word", "charts", "_rels", fmt.Sprintf("chart%d.xml.rels", chartIndex))
	target, err := findRelationshipTarget(relsPath, relID)
	if err != nil {
		return "", fmt.Errorf("resolve relationship %s for chart%d: %w", relID, chartIndex, err)
	}
	if target == "" {
		return "", fmt.Errorf("relationship %s for chart%d has empty target", relID, chartIndex)
	}

	// Relationship targets are relative to the source part (chart#.xml), not the .rels folder.
	resolved := filepath.Clean(filepath.Join(filepath.Dir(chartPath), filepath.FromSlash(target)))
	if _, statErr := os.Stat(resolved); statErr != nil {
		return "", fmt.Errorf("workbook file %s for chart%d not found: %w", resolved, chartIndex, statErr)
	}

	return resolved, nil
}

func externalDataRelID(chartXML []byte) string {
	content := string(chartXML)
	// Try both with and without namespace prefix
	markers := []string{"<c:externalData", "<externalData"}
	var tag string
	for _, marker := range markers {
		start := strings.Index(content, marker)
		if start == -1 {
			continue
		}
		end := strings.Index(content[start:], ">")
		if end == -1 {
			continue
		}
		tag = content[start : start+end]
		break
	}
	if tag == "" {
		return ""
	}
	// Try both r:id and relationships:id attribute names
	relAttrs := []string{`r:id="`, `relationships:id="`}
	for _, relAttr := range relAttrs {
		attrIdx := strings.Index(tag, relAttr)
		if attrIdx == -1 {
			continue
		}
		valueStart := attrIdx + len(relAttr)
		valueEnd := strings.Index(tag[valueStart:], `"`)
		if valueEnd == -1 {
			continue
		}
		return tag[valueStart : valueStart+valueEnd]
	}
	return ""
}

type relationships struct {
	XMLName       xml.Name       `xml:"Relationships"`
	Relationships []relationship `xml:"Relationship"`
}

type relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

func findRelationshipTarget(relsPath, relationshipID string) (string, error) {
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read relationships: %w", err)
	}
	var rels relationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return "", fmt.Errorf("parse relationships: %w", err)
	}
	for _, rel := range rels.Relationships {
		if rel.ID == relationshipID {
			return rel.Target, nil
		}
	}
	return "", fmt.Errorf("relationship %s not found", relationshipID)
}

// normalizeTemplateToDocument promotes a DOTX template content type to a DOCX
// document content type in [Content_Types].xml. It is a no-op for regular DOCX files.
func normalizeTemplateToDocument(tempDir string) error {
	ctPath := filepath.Join(tempDir, "[Content_Types].xml")
	data, err := os.ReadFile(ctPath)
	if err != nil {
		return fmt.Errorf("read [Content_Types].xml: %w", err)
	}
	content := string(data)
	if !strings.Contains(content, DotxMainContentType) {
		return nil
	}
	updated := strings.ReplaceAll(content, DotxMainContentType, DocxMainContentType)
	return atomicWriteFile(ctPath, []byte(updated), 0o644)
}

// normalizeNSPrefix rewrites content so that the given XML namespace is bound to
// newPrefix. It replaces the xmlns declaration and all occurrences of oldPrefix:
// with newPrefix:. Returns the updated content and whether a change was made.
// No-op when oldPrefix is empty (namespace not declared) or already equals newPrefix.
func normalizeNSPrefix(content, nsURI, oldPrefix, newPrefix string) (string, bool) {
	if oldPrefix == "" || oldPrefix == newPrefix {
		return content, false
	}
	oldDecl := `xmlns:` + oldPrefix + `="` + nsURI + `"`
	content = strings.ReplaceAll(content, oldDecl, `xmlns:`+newPrefix+`="`+nsURI+`"`)
	content = strings.ReplaceAll(content, oldPrefix+":", newPrefix+":")
	return content, true
}

// normalizeWMLNamespacePrefix rewrites word/document.xml so that the
// WordprocessingML namespace is bound to the canonical "w" prefix and the
// relationships namespace is bound to the canonical "r" prefix.
//
// Tools such as python-docx, lxml, and LibreOffice sometimes produce DOCX files
// where these namespaces are declared with generated prefixes (e.g. "ns0:", "ns3:")
// instead of the conventional "w:" / "r:". This library hardcodes both prefixes
// throughout (e.g. "</w:body>", "r:id="…"") so any non-canonical prefix must be
// normalised once at load time.
//
// The fix is a targeted string replacement performed once at load time:
//   - xmlns:OLDPREFIX="...wordprocessingml..." → xmlns:w="...wordprocessingml..."
//   - xmlns:OLDPREFIX="...relationships"       → xmlns:r="...relationships"
//   - Every OLDPREFIX: occurrence              → canonical:
//
// This is safe because XML namespace prefixes are local aliases — the document
// semantics are identical with any prefix, and the canonical aliases are what
// Microsoft Word and this library both expect.
func normalizeWMLNamespacePrefix(tempDir string) error {
	const (
		wmlNS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		relsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	)

	docPath := filepath.Join(tempDir, "word", "document.xml")
	data, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	content := string(data)

	content, wChanged := normalizeNSPrefix(content, wmlNS, detectWMLNamespacePrefix(content), "w")
	// Normalize relationships namespace prefix to canonical "r".
	// Some templates (e.g. generated by LibreOffice) declare this namespace with
	// a generated prefix such as "ns3" in the root element. This library
	// hardcodes "r:id" in header/footer references and other places, so an
	// undeclared "r:" prefix would produce invalid XML that Word / LibreOffice
	// refuses to open.
	content, rChanged := normalizeNSPrefix(content, relsNS, detectXMLNamespacePrefix(content, relsNS), "r")

	if !wChanged && !rChanged {
		return nil
	}
	return atomicWriteFile(docPath, []byte(content), 0o644)
}

// normalizeFontTableCharsets rewrites invalid w:characterSet values in fontTable.xml.
// LibreOffice emits IANA charset names (e.g. "utf-8", "windows-1252") but OOXML
// ST_UcharHexNumber requires 2-digit uppercase hex Windows charset IDs.
// Unknown IANA names are replaced with "00" (ANSI_CHARSET — safe fallback).
func normalizeFontTableCharsets(tempDir string) error {
	ftPath := filepath.Join(tempDir, "word", "fontTable.xml")
	data, err := os.ReadFile(ftPath)
	if os.IsNotExist(err) {
		return nil // no fontTable — nothing to normalize
	}
	if err != nil {
		return fmt.Errorf("read fontTable.xml: %w", err)
	}

	content := string(data)
	// LibreOffice adds a non-standard w:characterSet="..." attribute alongside the
	// already-valid w:val="XX" attribute on <w:charset> elements. The validator
	// rejects the IANA name values AND having two w:val attributes if we naively
	// replace it. The correct fix is to simply strip the extra attribute entirely:
	// the existing w:val already carries the correct Windows charset ID.
	reCharset := regexp.MustCompile(`\s+w:characterSet="[^"]*"`)
	normalized := reCharset.ReplaceAllString(content, "")

	if normalized == content {
		return nil
	}
	return atomicWriteFile(ftPath, []byte(normalized), 0o644)
}

// normalizeNumberingLvlJc rewrites invalid logical alignment values in
// numbering.xml level justification (<w:lvlJc>) to OOXML-compatible values.
func normalizeNumberingLvlJc(tempDir string) error {
	numberingPath := filepath.Join(tempDir, "word", "numbering.xml")
	data, err := os.ReadFile(numberingPath)
	if os.IsNotExist(err) {
		return nil
	}
	if err != nil {
		return fmt.Errorf("read numbering.xml: %w", err)
	}

	normalized := normalizeLvlJcValues(string(data))
	if normalized == string(data) {
		return nil
	}

	if err := atomicWriteFile(numberingPath, []byte(normalized), 0o644); err != nil {
		return fmt.Errorf("write numbering.xml: %w", err)
	}
	return nil
}

// validateStructure checks that required OpenXML parts exist.
func (u *Updater) validateStructure() error {
	required := []string{
		"word/document.xml",
		"word/_rels/document.xml.rels",
		"[Content_Types].xml",
	}
	for _, path := range required {
		fullPath := filepath.Join(u.tempDir, path)
		if _, err := os.Stat(fullPath); err != nil {
			return fmt.Errorf("missing required file %s", path)
		}
	}
	return nil
}

// writeBlankDocxStructure creates a minimal valid DOCX file structure in the given directory.
func writeBlankDocxStructure(dir string) error {
	files := map[string]string{
		"[Content_Types].xml":                               blankContentTypes,
		"_rels/.rels":                                       blankRels,
		filepath.Join("word", "document.xml"):               blankDocument,
		filepath.Join("word", "_rels", "document.xml.rels"): blankDocumentRels,
		filepath.Join("docProps", "core.xml"):               "", // generated below
		filepath.Join("docProps", "app.xml"):                blankAppXML,
	}

	now := time.Now().Format(time.RFC3339)
	files[filepath.Join("docProps", "core.xml")] = fmt.Sprintf(blankCoreXML, now, now)

	for relPath, content := range files {
		fullPath := filepath.Join(dir, relPath)
		if err := os.MkdirAll(filepath.Dir(fullPath), 0o755); err != nil {
			return fmt.Errorf("create dir for %s: %w", relPath, err)
		}
		if err := atomicWriteFile(fullPath, []byte(content), 0o644); err != nil {
			return fmt.Errorf("write %s: %w", relPath, err)
		}
	}

	return nil
}

const blankContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`

const blankRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

const blankDocument = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
<w:body>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>`

const blankDocumentRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`

const blankCoreXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<cp:revision>1</cp:revision>
<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>`

const blankAppXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>go-docx</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0000</AppVersion>
</Properties>`

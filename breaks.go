package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
)

// InsertPageBreak inserts a page break into the document
func (u *Updater) InsertPageBreak(opts BreakOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Generate page break XML
	pageBreakXML := generatePageBreakXML()

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Insert page break at the specified position
	updated, err := insertBreakAtPosition(raw, pageBreakXML, opts)
	if err != nil {
		return fmt.Errorf("insert page break: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// InsertSectionBreak inserts a section break into the document
func (u *Updater) InsertSectionBreak(opts BreakOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Default to next page if not specified
	if opts.SectionType == "" {
		opts.SectionType = SectionBreakNextPage
	}

	// Validate section break type
	if err := validateSectionBreakType(opts.SectionType); err != nil {
		return fmt.Errorf("invalid section break type: %w", err)
	}

	// Generate section break XML with optional page layout
	sectionBreakXML := generateSectionBreakXML(opts.SectionType, opts.PageLayout)

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Insert section break at the specified position
	updated, err := insertBreakAtPosition(raw, sectionBreakXML, opts)
	if err != nil {
		return fmt.Errorf("insert section break: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// generatePageBreakXML creates the XML for a page break
// A page break in Word is represented by a paragraph containing a run with a break element
func generatePageBreakXML() []byte {
	// Page break is a paragraph with a run containing a page break
	return []byte(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)
}

// generateSectionBreakXML creates the XML for a section break
// Section breaks are more complex and define how the next section starts
func generateSectionBreakXML(breakType SectionBreakType, pageLayout *PageLayoutOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr>")
	buf.WriteString("<w:sectPr>")

	// Add section type
	buf.WriteString(fmt.Sprintf(`<w:type w:val="%s"/>`, breakType))

	// Use provided page layout or defaults
	if pageLayout == nil {
		pageLayout = PageLayoutLetterPortrait()
	}

	// Page size with orientation
	orientAttr := ""
	if pageLayout.Orientation == OrientationLandscape {
		orientAttr = ` w:orient="landscape"`
	}
	buf.WriteString(fmt.Sprintf(`<w:pgSz w:w="%d" w:h="%d"%s/>`,
		pageLayout.PageWidth, pageLayout.PageHeight, orientAttr))

	// Page margins
	buf.WriteString(fmt.Sprintf(`<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="%d" w:footer="%d" w:gutter="%d"/>`,
		pageLayout.MarginTop,
		pageLayout.MarginRight,
		pageLayout.MarginBottom,
		pageLayout.MarginLeft,
		pageLayout.MarginHeader,
		pageLayout.MarginFooter,
		pageLayout.MarginGutter))

	// Default column count
	buf.WriteString(`<w:cols w:space="720"/>`)

	buf.WriteString("</w:sectPr>")
	buf.WriteString("</w:pPr>")
	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// validateSectionBreakType validates the section break type
func validateSectionBreakType(breakType SectionBreakType) error {
	switch breakType {
	case SectionBreakNextPage, SectionBreakContinuous, SectionBreakEvenPage, SectionBreakOddPage:
		return nil
	default:
		return fmt.Errorf("invalid section break type: %s", breakType)
	}
}

// insertBreakAtPosition inserts a break (page or section) at the specified position
func insertBreakAtPosition(raw []byte, breakXML []byte, opts BreakOptions) ([]byte, error) {

	switch opts.Position {
	case PositionBeginning:
		return insertAtBodyStart(raw, breakXML)

	case PositionEnd:
		return insertAtBodyEnd(raw, breakXML)

	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionAfterText")
		}
		return insertAfterText(raw, breakXML, opts.Anchor)

	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionBeforeText")
		}
		return insertBeforeText(raw, breakXML, opts.Anchor)

	default:
		return nil, fmt.Errorf("invalid position: %d", opts.Position)
	}
}

// insertBreakAfterAnchor inserts a break after the paragraph containing the anchor text
func insertBreakAfterAnchor(raw []byte, breakXML []byte, anchor string) ([]byte, error) {
	anchorBytes := []byte(anchor)
	pos := bytes.Index(raw, anchorBytes)
	if pos == -1 {
		return nil, fmt.Errorf("anchor text not found: %s", anchor)
	}

	// Find the end of the paragraph containing the anchor
	paraEnd := []byte("</w:p>")
	endPos := bytes.Index(raw[pos:], paraEnd)
	if endPos == -1 {
		return nil, fmt.Errorf("paragraph end not found after anchor text")
	}
	insertPos := pos + endPos + len(paraEnd)

	result := make([]byte, 0, len(raw)+len(breakXML))
	result = append(result, raw[:insertPos]...)
	result = append(result, breakXML...)
	result = append(result, raw[insertPos:]...)
	return result, nil
}

// insertBreakBeforeAnchor inserts a break before the paragraph containing the anchor text
func insertBreakBeforeAnchor(raw []byte, breakXML []byte, anchor string) ([]byte, error) {
	anchorBytes := []byte(anchor)
	before, _, ok := bytes.Cut(raw, anchorBytes)
	if !ok {
		return nil, fmt.Errorf("anchor text not found: %s", anchor)
	}

	// Find the start of the paragraph containing the anchor
	paraStart := []byte("<w:p>")
	// Search backwards from anchor position
	startPos := bytes.LastIndex(before, paraStart)
	if startPos == -1 {
		// Try with paragraph properties
		paraStart = []byte("<w:p ")
		startPos = bytes.LastIndex(before, paraStart)
		if startPos == -1 {
			return nil, fmt.Errorf("paragraph start not found before anchor text")
		}
	}

	result := make([]byte, 0, len(raw)+len(breakXML))
	result = append(result, raw[:startPos]...)
	result = append(result, breakXML...)
	result = append(result, raw[startPos:]...)
	return result, nil
}

// detectWMLNamespacePrefix returns the XML namespace prefix bound to the
// WordprocessingML namespace URI in the given document.xml content.
// Standard Word output uses "w"; LibreOffice and python-docx may serialize
// the same namespace with a generated prefix such as "ns0".
func detectWMLNamespacePrefix(content string) string {
	const wmlNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	re := regexp.MustCompile(`xmlns:([a-zA-Z][a-zA-Z0-9]*)="` + regexp.QuoteMeta(wmlNS) + `"`)
	if m := re.FindStringSubmatch(content); len(m) > 1 {
		return m[1]
	}
	return "w" // canonical default – always correct for MS Word documents
}

// SetPageLayout sets the page layout for the current or last section in the document.
// It detects the WordprocessingML namespace prefix actually used in the document so
// that it works correctly with files produced by MS Word ("w:"), LibreOffice ("ns0:"),
// python-docx, and other conforming producers. MS Word canonical output is preferred
// and the fallback when no explicit prefix mapping is found.
func (u *Updater) SetPageLayout(pageLayout PageLayoutOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	content := string(raw)

	// Detect the WordprocessingML namespace prefix actually used in this document.
	// Standard Word uses "w:", but tools like LibreOffice or python-docx may use "ns0:" or similar.
	ns := detectWMLNamespacePrefix(content)

	// Build tag strings for the detected prefix.
	// We search for both forms: <ns:sectPr> (no attributes) and <ns:sectPr (with attributes
	// such as w:rsidR="..."), which are both common in real-world documents.
	sectPrNoAttr := fmt.Sprintf("<%s:sectPr>", ns)
	sectPrWithAttr := fmt.Sprintf("<%s:sectPr ", ns)
	sectPrEnd := fmt.Sprintf("</%s:sectPr>", ns)
	bodyEnd := fmt.Sprintf("</%s:body>", ns)

	// Find the last sectPr opening tag, considering both forms.
	lastIdx := -1
	for _, tag := range []string{sectPrNoAttr, sectPrWithAttr} {
		sp := 0
		for {
			idx := bytes.Index([]byte(content[sp:]), []byte(tag))
			if idx == -1 {
				break
			}
			absIdx := sp + idx
			if absIdx > lastIdx {
				lastIdx = absIdx
			}
			sp = absIdx + 1
		}
	}

	if lastIdx == -1 {
		// No sectPr found – insert one immediately before </body>.
		bodyIdx := bytes.Index([]byte(content), []byte(bodyEnd))
		if bodyIdx == -1 {
			return fmt.Errorf("document body not found")
		}
		sectPrXML := generateSectionPropertiesXMLWithPrefix(ns, pageLayout)
		content = content[:bodyIdx] + sectPrXML + content[bodyIdx:]
	} else {
		// Replace the existing sectPr entirely.
		endIdx := bytes.Index([]byte(content[lastIdx:]), []byte(sectPrEnd))
		if endIdx == -1 {
			return fmt.Errorf("malformed section properties: closing tag not found")
		}
		endIdx += lastIdx + len(sectPrEnd)
		sectPrXML := generateSectionPropertiesXMLWithPrefix(ns, pageLayout)
		content = content[:lastIdx] + sectPrXML + content[endIdx:]
	}

	// Write updated document
	if err := os.WriteFile(docPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// generateSectionPropertiesXML generates section properties XML using the canonical "w:" prefix.
// This is the standard form produced by and expected by Microsoft Word.
func generateSectionPropertiesXML(pageLayout PageLayoutOptions) string {
	return generateSectionPropertiesXMLWithPrefix("w", pageLayout)
}

// generateSectionPropertiesXMLWithPrefix generates section properties XML using the given
// namespace prefix, allowing compatibility with documents produced by LibreOffice or other
// tools that may use a different prefix for the WordprocessingML namespace.
func generateSectionPropertiesXMLWithPrefix(ns string, pageLayout PageLayoutOptions) string {
	var buf bytes.Buffer

	fmt.Fprintf(&buf, "<%s:sectPr>", ns)

	// Page size with orientation
	orientAttr := ""
	if pageLayout.Orientation == OrientationLandscape {
		orientAttr = fmt.Sprintf(` %s:orient="landscape"`, ns)
	}
	fmt.Fprintf(&buf, `<%s:pgSz %s:w="%d" %s:h="%d"%s/>`,
		ns, ns, pageLayout.PageWidth, ns, pageLayout.PageHeight, orientAttr)

	// Page margins
	fmt.Fprintf(&buf, `<%s:pgMar %s:top="%d" %s:right="%d" %s:bottom="%d" %s:left="%d" %s:header="%d" %s:footer="%d" %s:gutter="%d"/>`,
		ns, ns, pageLayout.MarginTop,
		ns, pageLayout.MarginRight,
		ns, pageLayout.MarginBottom,
		ns, pageLayout.MarginLeft,
		ns, pageLayout.MarginHeader,
		ns, pageLayout.MarginFooter,
		ns, pageLayout.MarginGutter)

	// Default column settings
	fmt.Fprintf(&buf, `<%s:cols %s:space="720"/>`, ns, ns)

	fmt.Fprintf(&buf, "</%s:sectPr>", ns)

	return buf.String()
}

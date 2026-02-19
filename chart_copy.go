package docxupdater

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// CopyChart duplicates an existing chart and inserts it after the source chart's paragraph.
// sourceChartIndex: 1-based index of the chart to copy
// afterText is ignored (kept for backward-compatibility) and placement follows the source chart.
// Returns the new chart's index (1-based)
func (u *Updater) CopyChart(sourceChartIndex int, afterText string) (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}
	if sourceChartIndex < 1 {
		return 0, fmt.Errorf("source chart index must be >= 1")
	}

	// Find the next available chart index
	nextChartIndex := u.findNextChartIndex()

	// Copy chart XML file
	sourceChartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", sourceChartIndex))
	destChartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", nextChartIndex))
	if err := copyFile(sourceChartPath, destChartPath); err != nil {
		return 0, fmt.Errorf("copy chart xml: %w", err)
	}

	// Copy chart relationships file
	sourceRelsPath := filepath.Join(u.tempDir, "word", "charts", "_rels", fmt.Sprintf("chart%d.xml.rels", sourceChartIndex))
	destRelsPath := filepath.Join(u.tempDir, "word", "charts", "_rels", fmt.Sprintf("chart%d.xml.rels", nextChartIndex))
	if err := copyFile(sourceRelsPath, destRelsPath); err != nil {
		return 0, fmt.Errorf("copy chart relationships: %w", err)
	}

	// Copy embedded workbook
	sourceWorkbookPath, err := u.findWorkbookPathForChart(sourceChartIndex)
	if err != nil {
		return 0, fmt.Errorf("find source workbook: %w", err)
	}

	// Determine the new workbook filename
	destWorkbookPath := u.generateWorkbookPath(nextChartIndex, sourceWorkbookPath)
	if err := copyFile(sourceWorkbookPath, destWorkbookPath); err != nil {
		return 0, fmt.Errorf("copy workbook: %w", err)
	}

	// Update the chart relationships to point to the new workbook
	if err := u.updateChartRelationshipsForNewWorkbook(destRelsPath, sourceWorkbookPath, destWorkbookPath); err != nil {
		return 0, fmt.Errorf("update chart relationships: %w", err)
	}

	// Add chart relationship to document.xml.rels and get new rId
	newRelID, err := u.addChartRelationship(nextChartIndex)
	if err != nil {
		return 0, fmt.Errorf("add chart relationship: %w", err)
	}

	// Insert new drawing immediately after the source chart's paragraph
	if err := u.insertChartAfterSource(nextChartIndex, sourceChartIndex, newRelID); err != nil {
		return 0, fmt.Errorf("insert chart in document: %w", err)
	}

	// Update [Content_Types].xml
	if err := u.addContentTypeOverride(nextChartIndex); err != nil {
		return 0, fmt.Errorf("add content type: %w", err)
	}

	return nextChartIndex, nil
}

// findNextChartIndex finds the next available chart index
func (u *Updater) findNextChartIndex() int {
	chartsDir := filepath.Join(u.tempDir, "word", "charts")
	entries, err := os.ReadDir(chartsDir)
	if err != nil {
		return 1
	}

	maxIndex := 0
	for _, entry := range entries {
		if matches := chartFilePattern.FindStringSubmatch(entry.Name()); matches != nil {
			var idx int
			fmt.Sscanf(matches[1], "%d", &idx)
			if idx > maxIndex {
				maxIndex = idx
			}
		}
	}

	return maxIndex + 1
}

// generateWorkbookPath generates the destination path for the copied workbook
func (u *Updater) generateWorkbookPath(chartIndex int, sourceWorkbookPath string) string {
	// Extract the workbook filename pattern (e.g., embeddings/Microsoft_Excel_Worksheet1.xlsx)
	relPath, _ := filepath.Rel(filepath.Join(u.tempDir, "word", "charts"), sourceWorkbookPath)

	// If it has a numbered suffix, increment it; otherwise add chartIndex
	base := filepath.Base(relPath)
	ext := filepath.Ext(base)
	nameWithoutExt := strings.TrimSuffix(base, ext)

	// Try to find existing numeric suffix
	if matches := workbookNumberPattern.FindStringSubmatch(nameWithoutExt); matches != nil {
		// Has numeric suffix - use the chart index as the new suffix
		newName := fmt.Sprintf("%s%d%s", matches[1], chartIndex, ext)
		return filepath.Join(filepath.Dir(sourceWorkbookPath), newName)
	}

	// No numeric suffix - add chart index
	newName := fmt.Sprintf("%s%d%s", nameWithoutExt, chartIndex, ext)
	return filepath.Join(filepath.Dir(sourceWorkbookPath), newName)
}

// updateChartRelationshipsForNewWorkbook updates the relationship target in the chart rels file
func (u *Updater) updateChartRelationshipsForNewWorkbook(relsPath, oldWorkbookPath, newWorkbookPath string) error {
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read relationships: %w", err)
	}

	var rels relationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return fmt.Errorf("parse relationships: %w", err)
	}

	// Get relative paths from charts directory
	chartsDir := filepath.Join(u.tempDir, "word", "charts")
	oldRelPath, _ := filepath.Rel(chartsDir, oldWorkbookPath)
	newRelPath, _ := filepath.Rel(chartsDir, newWorkbookPath)

	// Update the target for the embedded workbook relationship
	for i := range rels.Relationships {
		if strings.Contains(rels.Relationships[i].Target, oldRelPath) {
			rels.Relationships[i].Target = strings.Replace(
				rels.Relationships[i].Target,
				oldRelPath,
				newRelPath,
				1,
			)
		}
	}

	// Marshal back to XML
	output, err := xml.MarshalIndent(rels, "", "  ")
	if err != nil {
		return fmt.Errorf("marshal relationships: %w", err)
	}

	// Add XML header
	result := []byte(xml.Header)
	result = append(result, output...)

	if err := os.WriteFile(relsPath, result, 0o644); err != nil {
		return fmt.Errorf("write relationships: %w", err)
	}

	return nil
}

// insertChartAfterSource adds a chart drawing after the paragraph that contains the source chart
func (u *Updater) insertChartAfterSource(newChartIndex int, sourceChartIndex int, newRelId string) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Locate the relationship Id for the source chart (e.g., rId5 for charts/chart1.xml)
	sourceRelID, err := u.getRelIdForChart(sourceChartIndex)
	if err != nil {
		return err
	}
	// Find the drawing that references this r:id in document.xml
	tag := fmt.Appendf(nil, "r:id=\"%s\"", sourceRelID)
	drawIdx := bytes.Index(raw, tag)
	if drawIdx == -1 {
		return fmt.Errorf("could not find drawing for source chart %d", sourceChartIndex)
	}
	// Find the end of the parent paragraph </w:p> after this drawing
	closingTag := []byte("</w:p>")
	paraEndRel := bytes.Index(raw[drawIdx:], closingTag)
	if paraEndRel == -1 {
		return fmt.Errorf("could not find end of source paragraph")
	}
	insertPos := drawIdx + paraEndRel + len(closingTag)

	// Generate chart drawing XML for the new chart using its new relationship id
	chartDrawing, err := u.generateChartDrawingXML(newChartIndex, newRelId)
	if err != nil {
		return fmt.Errorf("generate drawing xml: %w", err)
	}

	// Insert the chart drawing after the paragraph (pre-allocate for efficiency)
	result := make([]byte, len(raw)+len(chartDrawing))
	n := copy(result, raw[:insertPos])
	n += copy(result[n:], chartDrawing)
	copy(result[n:], raw[insertPos:])

	if err := os.WriteFile(docPath, result, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// getRelIdForChart finds the document relationship Id for charts/chart{index}.xml
func (u *Updater) getRelIdForChart(chartIndex int) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document relationships: %w", err)
	}
	// Build pattern for this specific chart
	pattern := regexp.MustCompile(fmt.Sprintf(chartRelPatternTemplate, chartIndex))
	m := pattern.FindSubmatch(raw)
	if m == nil {
		return "", fmt.Errorf("relationship for chart%d.xml not found", chartIndex)
	}
	return string(m[1]), nil
}

// getNextDocPrId finds the next available docPr ID in the document
func (u *Updater) getNextDocPrId() (int, error) {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document: %w", err)
	}

	// Find all docPr id values using package-level regex
	matches := docPrIDPattern.FindAllStringSubmatch(string(raw), -1)

	maxId := 0
	for _, match := range matches {
		if len(match) > 1 {
			var id int
			fmt.Sscanf(match[1], "%d", &id)
			if id > maxId {
				maxId = id
			}
		}
	}

	return maxId + 1, nil
}

// getNextDocumentRelId finds the next available relationship ID in document.xml.rels
func (u *Updater) getNextDocumentRelId() (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document rels: %w", err)
	}

	var rels relationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return "", fmt.Errorf("parse document rels: %w", err)
	}

	maxId := 0
	for _, rel := range rels.Relationships {
		if matches := relIDPattern.FindStringSubmatch(rel.ID); matches != nil {
			var id int
			fmt.Sscanf(matches[1], "%d", &id)
			if id > maxId {
				maxId = id
			}
		}
	}

	return fmt.Sprintf("rId%d", maxId+1), nil
}

// generateChartDrawingXML creates the inline drawing XML for a chart
func (u *Updater) generateChartDrawingXML(chartIndex int, relId string) ([]byte, error) {
	// Get a unique docPr ID (document-wide drawing object ID)
	docPrId, err := u.getNextDocPrId()
	if err != nil {
		return nil, fmt.Errorf("get next docPr id: %w", err)
	}

	// This generates a chart drawing paragraph matching Word/LibreOffice structure
	// Note: wp14 namespace is declared in the document root
	const template = `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="%08X" wp14:editId="%08X"><wp:extent cx="6099523" cy="3340467"/><wp:effectExtent l="0" t="0" r="15875" b="12700"/><wp:docPr id="%d" name="Chart %d"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

	// Generate unique IDs using constants
	anchorId := ChartAnchorIDBase + uint32(chartIndex)*ChartIDIncrement
	editId := ChartEditIDBase + uint32(chartIndex)*ChartIDIncrement

	return fmt.Appendf(nil, template, anchorId, editId, docPrId, chartIndex, relId), nil
}

// addChartRelationship appends a Relationship for the new chart to document.xml.rels and returns its Id
func (u *Updater) addChartRelationship(chartIndex int) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document relationships: %w", err)
	}

	// Compute next Id by scanning existing content
	nextRelId, err := u.getNextDocumentRelId()
	if err != nil {
		return "", err
	}

	// Append a self-closing Relationship before </Relationships>
	insert := fmt.Sprintf("\n  <Relationship Id=\"%s\" Type=\"%s/chart\" Target=\"charts/chart%d.xml\"/>\n", nextRelId, OfficeDocumentNS, chartIndex)
	closer := []byte("</Relationships>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return "", fmt.Errorf("invalid document.xml.rels: missing </Relationships>")
	}
	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])

	if err := os.WriteFile(relsPath, result, 0o644); err != nil {
		return "", fmt.Errorf("write relationships: %w", err)
	}
	return nextRelId, nil
}

// addContentTypeOverride adds a content type override for the new chart in [Content_Types].xml
func (u *Updater) addContentTypeOverride(chartIndex int) error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	raw, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	chartPart := fmt.Sprintf("/word/charts/chart%d.xml", chartIndex)
	if bytes.Contains(raw, []byte(chartPart)) {
		return nil // already present
	}

	insert := fmt.Sprintf("\n  <Override PartName=\"%s\" ContentType=\"%s\"/>\n", chartPart, ChartContentType)
	closer := []byte("</Types>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return fmt.Errorf("invalid [Content_Types].xml: missing </Types>")
	}
	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])
	return os.WriteFile(contentTypesPath, result, 0o644)
}

// copyFile copies a file from src to dst
func copyFile(src, dst string) error {
	sourceFile, err := os.Open(src)
	if err != nil {
		return fmt.Errorf("open source: %w", err)
	}
	defer sourceFile.Close()

	if err := os.MkdirAll(filepath.Dir(dst), 0o755); err != nil {
		return fmt.Errorf("create destination directory: %w", err)
	}

	destFile, err := os.Create(dst)
	if err != nil {
		return fmt.Errorf("create destination: %w", err)
	}
	defer destFile.Close()

	if _, err := io.Copy(destFile, sourceFile); err != nil {
		return fmt.Errorf("copy data: %w", err)
	}

	return nil
}

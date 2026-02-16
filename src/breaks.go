package docxupdater

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
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

	// Generate section break XML
	sectionBreakXML := generateSectionBreakXML(opts.SectionType)

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
func generateSectionBreakXML(breakType SectionBreakType) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr>")
	buf.WriteString("<w:sectPr>")

	// Add section type
	buf.WriteString(fmt.Sprintf(`<w:type w:val="%s"/>`, breakType))

	// Default page settings (8.5" x 11" letter size)
	// Width: 12240 twips (8.5"), Height: 15840 twips (11")
	buf.WriteString(`<w:pgSz w:w="12240" w:h="15840"/>`)

	// Default margins (1" all around = 1440 twips)
	buf.WriteString(`<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`)

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
	bodyStart := []byte("<w:body>")
	bodyEnd := []byte("</w:body>")

	startIdx := bytes.Index(raw, bodyStart)
	endIdx := bytes.LastIndex(raw, bodyEnd)
	if startIdx == -1 || endIdx == -1 {
		return nil, fmt.Errorf("invalid document.xml: missing <w:body>")
	}
	startIdx += len(bodyStart)

	switch opts.Position {
	case PositionBeginning:
		// Insert at the beginning of body
		result := make([]byte, 0, len(raw)+len(breakXML))
		result = append(result, raw[:startIdx]...)
		result = append(result, breakXML...)
		result = append(result, raw[startIdx:]...)
		return result, nil

	case PositionEnd:
		// Insert at the end of body (before </w:body>)
		result := make([]byte, 0, len(raw)+len(breakXML))
		result = append(result, raw[:endIdx]...)
		result = append(result, breakXML...)
		result = append(result, raw[endIdx:]...)
		return result, nil

	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionAfterText")
		}
		// Find the anchor text and insert after its paragraph
		return insertBreakAfterAnchor(raw, breakXML, opts.Anchor)

	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionBeforeText")
		}
		// Find the anchor text and insert before its paragraph
		return insertBreakBeforeAnchor(raw, breakXML, opts.Anchor)

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
	pos := bytes.Index(raw, anchorBytes)
	if pos == -1 {
		return nil, fmt.Errorf("anchor text not found: %s", anchor)
	}

	// Find the start of the paragraph containing the anchor
	paraStart := []byte("<w:p>")
	// Search backwards from anchor position
	startPos := bytes.LastIndex(raw[:pos], paraStart)
	if startPos == -1 {
		// Try with paragraph properties
		paraStart = []byte("<w:p ")
		startPos = bytes.LastIndex(raw[:pos], paraStart)
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

package docxupdater

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// HeaderType defines the type of header
type HeaderType string

const (
	// HeaderFirst is the header for the first page
	HeaderFirst HeaderType = "first"
	// HeaderEven is the header for even pages
	HeaderEven HeaderType = "even"
	// HeaderDefault is the default header (odd pages)
	HeaderDefault HeaderType = "default"
)

// FooterType defines the type of footer
type FooterType string

const (
	// FooterFirst is the footer for the first page
	FooterFirst FooterType = "first"
	// FooterEven is the footer for even pages
	FooterEven FooterType = "even"
	// FooterDefault is the default footer (odd pages)
	FooterDefault FooterType = "default"
)

// HeaderFooterContent defines the content structure for headers/footers
type HeaderFooterContent struct {
	// LeftText is left-aligned text
	LeftText string

	// CenterText is center-aligned text
	CenterText string

	// RightText is right-aligned text
	RightText string

	// PageNumber includes page number field
	PageNumber bool

	// PageNumberFormat defines the page number format (e.g., "Page X of Y")
	PageNumberFormat string

	// Date includes current date field
	Date bool

	// DateFormat defines date format
	DateFormat string
}

// HeaderOptions defines options for header creation
type HeaderOptions struct {
	// Type of header (first, even, default)
	Type HeaderType

	// DifferentFirst enables different header on first page
	DifferentFirst bool

	// DifferentOddEven enables different headers for odd/even pages
	DifferentOddEven bool
}

// FooterOptions defines options for footer creation
type FooterOptions struct {
	// Type of footer (first, even, default)
	Type FooterType

	// DifferentFirst enables different footer on first page
	DifferentFirst bool

	// DifferentOddEven enables different footers for odd/even pages
	DifferentOddEven bool
}

// DefaultHeaderOptions returns header options with sensible defaults
func DefaultHeaderOptions() HeaderOptions {
	return HeaderOptions{
		Type:             HeaderDefault,
		DifferentFirst:   false,
		DifferentOddEven: false,
	}
}

// DefaultFooterOptions returns footer options with sensible defaults
func DefaultFooterOptions() FooterOptions {
	return FooterOptions{
		Type:             FooterDefault,
		DifferentFirst:   false,
		DifferentOddEven: false,
	}
}

// SetHeader sets or creates a header for the document
func (u *Updater) SetHeader(content HeaderFooterContent, opts HeaderOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Determine header filename based on type
	var headerFile string
	switch opts.Type {
	case HeaderFirst:
		headerFile = "header1.xml"
	case HeaderEven:
		headerFile = "header2.xml"
	case HeaderDefault:
		headerFile = "header3.xml"
	default:
		headerFile = "header.xml"
	}

	headerPath := filepath.Join(u.tempDir, "word", headerFile)

	// Generate header XML
	headerXML := u.generateHeaderFooterXML(content, true)

	// Write header file
	if err := os.WriteFile(headerPath, headerXML, 0o644); err != nil {
		return NewHeaderFooterError("failed to write header", err)
	}

	// Add header relationship and get the relationship ID
	relID, err := u.addHeaderFooterRelationship(headerFile, "header")
	if err != nil {
		return NewHeaderFooterError("failed to add header relationship", err)
	}

	// Update document.xml to reference header
	if err := u.updateDocumentForHeaderFooter(opts.Type, "header", relID, opts.DifferentFirst, opts.DifferentOddEven); err != nil {
		return NewHeaderFooterError("failed to update document", err)
	}

	// Add content type for header
	if err := u.addHeaderFooterContentType(headerFile, "header"); err != nil {
		return NewHeaderFooterError("failed to add content type", err)
	}

	return nil
}

// SetFooter sets or creates a footer for the document
func (u *Updater) SetFooter(content HeaderFooterContent, opts FooterOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Determine footer filename based on type
	var footerFile string
	switch opts.Type {
	case FooterFirst:
		footerFile = "footer1.xml"
	case FooterEven:
		footerFile = "footer2.xml"
	case FooterDefault:
		footerFile = "footer3.xml"
	default:
		footerFile = "footer.xml"
	}

	footerPath := filepath.Join(u.tempDir, "word", footerFile)

	// Generate footer XML
	footerXML := u.generateHeaderFooterXML(content, false)

	// Write footer file
	if err := os.WriteFile(footerPath, footerXML, 0o644); err != nil {
		return NewHeaderFooterError("failed to write footer", err)
	}

	// Add footer relationship and get the relationship ID
	relID, err := u.addHeaderFooterRelationship(footerFile, "footer")
	if err != nil {
		return NewHeaderFooterError("failed to add footer relationship", err)
	}

	// Update document.xml to reference footer
	if err := u.updateDocumentForHeaderFooter(opts.Type, "footer", relID, opts.DifferentFirst, opts.DifferentOddEven); err != nil {
		return NewHeaderFooterError("failed to update document", err)
	}

	// Add content type for footer
	if err := u.addHeaderFooterContentType(footerFile, "footer"); err != nil {
		return NewHeaderFooterError("failed to add content type", err)
	}

	return nil
}

// generateHeaderFooterXML creates the XML content for a header or footer
func (u *Updater) generateHeaderFooterXML(content HeaderFooterContent, isHeader bool) []byte {
	var buf strings.Builder

	rootElement := "w:hdr"
	if !isHeader {
		rootElement = "w:ftr"
	}

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(fmt.Sprintf(`<%s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `, rootElement))
	buf.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString("\n")

	// Create table for left, center, right layout
	if content.LeftText != "" || content.CenterText != "" || content.RightText != "" {
		buf.WriteString(u.generateThreeColumnTable(content))
	}

	// Add page number if requested
	if content.PageNumber {
		buf.WriteString(u.generatePageNumberParagraph(content.PageNumberFormat))
	}

	// Add date if requested
	if content.Date {
		buf.WriteString(u.generateDateParagraph(content.DateFormat))
	}

	buf.WriteString(fmt.Sprintf("</%s>", rootElement))

	return []byte(buf.String())
}

// generateThreeColumnTable creates a table with left, center, right columns
func (u *Updater) generateThreeColumnTable(content HeaderFooterContent) string {
	var buf strings.Builder

	buf.WriteString("<w:tbl>")
	buf.WriteString("<w:tblPr>")
	buf.WriteString(`<w:tblW w:w="5000" w:type="pct"/>`) // 100% width
	buf.WriteString(`<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>`)
	buf.WriteString("</w:tblPr>")
	buf.WriteString("<w:tblGrid>")
	buf.WriteString(`<w:gridCol w:w="3000"/>`) // Left column
	buf.WriteString(`<w:gridCol w:w="3000"/>`) // Center column
	buf.WriteString(`<w:gridCol w:w="3000"/>`) // Right column
	buf.WriteString("</w:tblGrid>")

	// Single row
	buf.WriteString("<w:tr>")

	// Left cell
	buf.WriteString("<w:tc>")
	buf.WriteString("<w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>")
	if content.LeftText != "" {
		buf.WriteString("<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr>")
		buf.WriteString(fmt.Sprintf("<w:r><w:t>%s</w:t></w:r>", escapeXML(content.LeftText)))
		buf.WriteString("</w:p>")
	} else {
		buf.WriteString("<w:p/>")
	}
	buf.WriteString("</w:tc>")

	// Center cell
	buf.WriteString("<w:tc>")
	buf.WriteString("<w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>")
	if content.CenterText != "" {
		buf.WriteString("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>")
		buf.WriteString(fmt.Sprintf("<w:r><w:t>%s</w:t></w:r>", escapeXML(content.CenterText)))
		buf.WriteString("</w:p>")
	} else {
		buf.WriteString("<w:p/>")
	}
	buf.WriteString("</w:tc>")

	// Right cell
	buf.WriteString("<w:tc>")
	buf.WriteString("<w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>")
	if content.RightText != "" {
		buf.WriteString("<w:p><w:pPr><w:jc w:val=\"right\"/></w:pPr>")
		buf.WriteString(fmt.Sprintf("<w:r><w:t>%s</w:t></w:r>", escapeXML(content.RightText)))
		buf.WriteString("</w:p>")
	} else {
		buf.WriteString("<w:p/>")
	}
	buf.WriteString("</w:tc>")

	buf.WriteString("</w:tr>")
	buf.WriteString("</w:tbl>")

	return buf.String()
}

// generatePageNumberParagraph creates a paragraph with page number field
func (u *Updater) generatePageNumberParagraph(format string) string {
	var buf strings.Builder

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr><w:jc w:val=\"center\"/></w:pPr>")
	buf.WriteString("<w:r>")

	if format != "" {
		// Custom format (e.g., "Page X of Y")
		parts := strings.Split(format, "X")
		if len(parts) > 0 && parts[0] != "" {
			buf.WriteString(fmt.Sprintf("<w:t>%s</w:t></w:r><w:r>", escapeXML(parts[0])))
		}
	}

	// Page number field
	buf.WriteString("<w:fldChar w:fldCharType=\"begin\"/>")
	buf.WriteString("</w:r><w:r>")
	buf.WriteString("<w:instrText>PAGE</w:instrText>")
	buf.WriteString("</w:r><w:r>")
	buf.WriteString("<w:fldChar w:fldCharType=\"end\"/>")

	if format != "" && strings.Contains(format, "Y") {
		// Add " of Y" (total pages)
		buf.WriteString("</w:r><w:r>")
		buf.WriteString("<w:t> of </w:t>")
		buf.WriteString("</w:r><w:r>")
		buf.WriteString("<w:fldChar w:fldCharType=\"begin\"/>")
		buf.WriteString("</w:r><w:r>")
		buf.WriteString("<w:instrText>NUMPAGES</w:instrText>")
		buf.WriteString("</w:r><w:r>")
		buf.WriteString("<w:fldChar w:fldCharType=\"end\"/>")
	}

	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	return buf.String()
}

// generateDateParagraph creates a paragraph with date field
func (u *Updater) generateDateParagraph(format string) string {
	var buf strings.Builder

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr><w:jc w:val=\"right\"/></w:pPr>")
	buf.WriteString("<w:r>")
	buf.WriteString("<w:fldChar w:fldCharType=\"begin\"/>")
	buf.WriteString("</w:r><w:r>")

	dateFormat := format
	if dateFormat == "" {
		dateFormat = "MMMM d, yyyy" // Default: January 1, 2026
	}

	buf.WriteString(fmt.Sprintf("<w:instrText>DATE \\@ \"%s\"</w:instrText>", dateFormat))
	buf.WriteString("</w:r><w:r>")
	buf.WriteString("<w:fldChar w:fldCharType=\"end\"/>")
	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	return buf.String()
}

// addHeaderFooterRelationship adds a relationship for header/footer and returns the relationship ID
func (u *Updater) addHeaderFooterRelationship(filename, hdrFtrType string) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")

	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read relationships: %w", err)
	}

	content := string(raw)

	// Check if relationship already exists
	existingIDPattern := fmt.Sprintf(`<Relationship Id="([^"]+)"[^>]*Target="%s"`, regexp.QuoteMeta(filename))
	if match := regexp.MustCompile(existingIDPattern).FindStringSubmatch(content); len(match) > 1 {
		return match[1], nil // Already exists, return existing ID
	}

	// Find next available relationship ID
	nextID := u.getNextRelationshipID(content)
	relID := fmt.Sprintf("rId%d", nextID)

	// Create header/footer relationship
	relType := "header"
	if hdrFtrType == "footer" {
		relType = "footer"
	}

	newRel := fmt.Sprintf(
		`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s" Target="%s"/>`,
		relID,
		relType,
		filename,
	)

	// Insert before closing </Relationships>
	content = strings.Replace(content, "</Relationships>", newRel+"</Relationships>", 1)

	// Write updated relationships
	if err := os.WriteFile(relsPath, []byte(content), 0o644); err != nil {
		return "", fmt.Errorf("write relationships: %w", err)
	}

	return relID, nil
}

// updateDocumentForHeaderFooter updates document.xml to reference header/footer
func (u *Updater) updateDocumentForHeaderFooter(hdrFtrType interface{}, hdrFtr string, relID string, differentFirst, differentOddEven bool) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")

	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document: %w", err)
	}

	content := string(raw)

	// Find or create <w:sectPr> section
	sectPrPattern := `<w:sectPr[^>]*>.*?</w:sectPr>`
	sectPrRegex := regexp.MustCompile(sectPrPattern)

	if sectPrRegex.MatchString(content) {
		// Update existing sectPr
		content = sectPrRegex.ReplaceAllStringFunc(content, func(sectPr string) string {
			return u.addHeaderFooterToSectPr(sectPr, hdrFtrType, hdrFtr, relID, differentFirst, differentOddEven)
		})
	} else {
		// Create new sectPr before </w:body>
		sectPr := u.createSectPrWithHeaderFooter(hdrFtrType, hdrFtr, relID, differentFirst, differentOddEven)
		content = strings.Replace(content, "</w:body>", sectPr+"</w:body>", 1)
	}

	// Write updated document
	if err := os.WriteFile(docPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write document: %w", err)
	}

	return nil
}

// addHeaderFooterToSectPr adds header/footer reference to existing sectPr
func (u *Updater) addHeaderFooterToSectPr(sectPr string, hdrFtrType interface{}, hdrFtr string, relID string, differentFirst, differentOddEven bool) string {
	// Determine reference type
	refType := "default"
	if hdrFtrType == HeaderFirst || hdrFtrType == FooterFirst {
		refType = "first"
	} else if hdrFtrType == HeaderEven || hdrFtrType == FooterEven {
		refType = "even"
	}

	// Create reference element
	var refElement string
	if hdrFtr == "header" {
		refElement = fmt.Sprintf(`<w:headerReference w:type="%s" r:id="%s"/>`, refType, relID)
	} else {
		refElement = fmt.Sprintf(`<w:footerReference w:type="%s" r:id="%s"/>`, refType, relID)
	}

	// Check if reference already exists for this type
	refPattern := fmt.Sprintf(`<w:%sReference w:type="%s"[^>]*/>`, hdrFtr, refType)
	if matched, _ := regexp.MatchString(refPattern, sectPr); matched {
		// Replace existing reference
		sectPr = regexp.MustCompile(refPattern).ReplaceAllString(sectPr, refElement)
	} else {
		// Insert before </w:sectPr>
		sectPr = strings.Replace(sectPr, "</w:sectPr>", refElement+"</w:sectPr>", 1)
	}

	return sectPr
}

// createSectPrWithHeaderFooter creates a new sectPr with header/footer
func (u *Updater) createSectPrWithHeaderFooter(hdrFtrType interface{}, hdrFtr string, relID string, differentFirst, differentOddEven bool) string {
	var buf strings.Builder

	buf.WriteString("<w:sectPr>")

	if differentFirst {
		buf.WriteString("<w:titlePg/>")
	}

	// Determine reference type
	refType := "default"
	if hdrFtrType == HeaderFirst || hdrFtrType == FooterFirst {
		refType = "first"
	} else if hdrFtrType == HeaderEven || hdrFtrType == FooterEven {
		refType = "even"
	}

	// Add header/footer reference with actual relationship ID
	if hdrFtr == "header" {
		buf.WriteString(fmt.Sprintf(`<w:headerReference w:type="%s" r:id="%s"/>`, refType, relID))
	} else {
		buf.WriteString(fmt.Sprintf(`<w:footerReference w:type="%s" r:id="%s"/>`, refType, relID))
	}

	buf.WriteString("</w:sectPr>")

	return buf.String()
}

// addHeaderFooterContentType adds content type for header/footer
func (u *Updater) addHeaderFooterContentType(filename, hdrFtrType string) error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")

	raw, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	content := string(raw)

	// Check if already exists
	if strings.Contains(content, filename) {
		return nil
	}

	contentType := "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	if hdrFtrType == "footer" {
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	}

	override := fmt.Sprintf(
		`<Override PartName="/word/%s" ContentType="%s"/>`,
		filename,
		contentType,
	)

	// Insert before closing </Types>
	content = strings.Replace(content, "</Types>", override+"</Types>", 1)

	// Write updated content types
	if err := os.WriteFile(contentTypesPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write content types: %w", err)
	}

	return nil
}

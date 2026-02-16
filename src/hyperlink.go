package docxupdater

import (
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// HyperlinkOptions defines options for hyperlink insertion
type HyperlinkOptions struct {
	// Position where to insert the hyperlink
	Position InsertPosition

	// Anchor text for position-based insertion (for PositionAfterText/PositionBeforeText)
	Anchor string

	// Tooltip text shown on hover
	Tooltip string

	// Style to apply to the hyperlink paragraph (if creating new paragraph)
	Style ParagraphStyle

	// Color for hyperlink text (hex color, e.g., "0563C1")
	Color string

	// Underline the hyperlink text
	Underline bool

	// ScreenTip for accessibility
	ScreenTip string
}

// DefaultHyperlinkOptions returns hyperlink options with sensible defaults
func DefaultHyperlinkOptions() HyperlinkOptions {
	return HyperlinkOptions{
		Position:  PositionEnd,
		Color:     "0563C1", // Standard Word hyperlink blue
		Underline: true,
		Style:     StyleNormal,
	}
}

// InsertHyperlink inserts a hyperlink into the document
func (u *Updater) InsertHyperlink(text, urlStr string, opts HyperlinkOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if text == "" {
		return NewValidationError("text", "hyperlink text cannot be empty")
	}
	if urlStr == "" {
		return NewValidationError("url", "hyperlink URL cannot be empty")
	}

	// Validate URL format
	if err := validateURL(urlStr); err != nil {
		return err
	}

	// Apply defaults
	if opts.Color == "" {
		opts.Color = "0563C1"
	}
	if opts.Style == "" {
		opts.Style = StyleNormal
	}

	// Add hyperlink relationship
	relID, err := u.addHyperlinkRelationship(urlStr)
	if err != nil {
		return NewHyperlinkError("failed to add relationship", err)
	}

	// Generate hyperlink XML
	hyperlinkXML := u.generateHyperlinkXML(text, relID, opts)

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return NewXMLParseError("document.xml", err)
	}

	// Insert hyperlink at specified position
	updated, err := u.insertHyperlinkAtPosition(raw, hyperlinkXML, opts)
	if err != nil {
		return fmt.Errorf("insert hyperlink: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return NewXMLWriteError("document.xml", err)
	}

	return nil
}

// InsertInternalLink inserts a link to a bookmark within the document
func (u *Updater) InsertInternalLink(text, bookmarkName string, opts HyperlinkOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if text == "" {
		return NewValidationError("text", "link text cannot be empty")
	}
	if bookmarkName == "" {
		return NewValidationError("bookmarkName", "bookmark name cannot be empty")
	}

	// Apply defaults
	if opts.Color == "" {
		opts.Color = "0563C1"
	}
	if opts.Style == "" {
		opts.Style = StyleNormal
	}

	// Generate internal hyperlink XML (uses anchor instead of rId)
	hyperlinkXML := u.generateInternalHyperlinkXML(text, bookmarkName, opts)

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return NewXMLParseError("document.xml", err)
	}

	// Insert hyperlink at specified position
	updated, err := u.insertHyperlinkAtPosition(raw, hyperlinkXML, opts)
	if err != nil {
		return fmt.Errorf("insert internal link: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return NewXMLWriteError("document.xml", err)
	}

	return nil
}

// addHyperlinkRelationship adds a hyperlink relationship to document.xml.rels
func (u *Updater) addHyperlinkRelationship(urlStr string) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")

	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read relationships: %w", err)
	}

	content := string(raw)

	// Find next available relationship ID
	nextID := u.getNextRelationshipID(content)
	relID := fmt.Sprintf("rId%d", nextID)

	// Create hyperlink relationship
	newRel := fmt.Sprintf(
		`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>`,
		relID,
		escapeXMLAttribute(urlStr),
	)

	// Insert before closing </Relationships>
	content = strings.Replace(content, "</Relationships>", newRel+"</Relationships>", 1)

	// Write updated relationships
	if err := os.WriteFile(relsPath, []byte(content), 0o644); err != nil {
		return "", fmt.Errorf("write relationships: %w", err)
	}

	return relID, nil
}

// getNextRelationshipID finds the next available relationship ID
func (u *Updater) getNextRelationshipID(relsContent string) int {
	pattern := regexp.MustCompile(`Id="rId(\d+)"`)
	matches := pattern.FindAllStringSubmatch(relsContent, -1)

	maxID := 0
	for _, match := range matches {
		if len(match) > 1 {
			var id int
			fmt.Sscanf(match[1], "%d", &id)
			if id > maxID {
				maxID = id
			}
		}
	}

	return maxID + 1
}

// generateHyperlinkXML creates the XML for a hyperlink
func (u *Updater) generateHyperlinkXML(text, relID string, opts HyperlinkOptions) []byte {
	var buf strings.Builder

	buf.WriteString("<w:p>")

	// Add paragraph properties
	if opts.Style != "" {
		buf.WriteString("<w:pPr>")
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, opts.Style))
		buf.WriteString("</w:pPr>")
	}

	// Start hyperlink
	buf.WriteString(fmt.Sprintf(`<w:hyperlink r:id="%s"`, relID))
	if opts.Tooltip != "" || opts.ScreenTip != "" {
		tooltip := opts.Tooltip
		if tooltip == "" {
			tooltip = opts.ScreenTip
		}
		buf.WriteString(fmt.Sprintf(` w:tooltip="%s"`, escapeXMLAttribute(tooltip)))
	}
	buf.WriteString(">")

	// Add run with text
	buf.WriteString("<w:r>")
	buf.WriteString("<w:rPr>")

	// Add hyperlink style
	buf.WriteString(`<w:rStyle w:val="Hyperlink"/>`)

	// Add color
	if opts.Color != "" {
		buf.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, opts.Color))
	}

	// Add underline
	if opts.Underline {
		buf.WriteString(`<w:u w:val="single"/>`)
	}

	buf.WriteString("</w:rPr>")
	buf.WriteString(fmt.Sprintf("<w:t>%s</w:t>", escapeXML(text)))
	buf.WriteString("</w:r>")

	buf.WriteString("</w:hyperlink>")
	buf.WriteString("</w:p>")

	return []byte(buf.String())
}

// generateInternalHyperlinkXML creates XML for internal document links
func (u *Updater) generateInternalHyperlinkXML(text, bookmarkName string, opts HyperlinkOptions) []byte {
	var buf strings.Builder

	buf.WriteString("<w:p>")

	// Add paragraph properties
	if opts.Style != "" {
		buf.WriteString("<w:pPr>")
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, opts.Style))
		buf.WriteString("</w:pPr>")
	}

	// Start hyperlink with anchor (internal link)
	buf.WriteString(fmt.Sprintf(`<w:hyperlink w:anchor="%s"`, escapeXMLAttribute(bookmarkName)))
	if opts.Tooltip != "" || opts.ScreenTip != "" {
		tooltip := opts.Tooltip
		if tooltip == "" {
			tooltip = opts.ScreenTip
		}
		buf.WriteString(fmt.Sprintf(` w:tooltip="%s"`, escapeXMLAttribute(tooltip)))
	}
	buf.WriteString(">")

	// Add run with text
	buf.WriteString("<w:r>")
	buf.WriteString("<w:rPr>")

	// Add hyperlink style
	buf.WriteString(`<w:rStyle w:val="Hyperlink"/>`)

	// Add color
	if opts.Color != "" {
		buf.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, opts.Color))
	}

	// Add underline
	if opts.Underline {
		buf.WriteString(`<w:u w:val="single"/>`)
	}

	buf.WriteString("</w:rPr>")
	buf.WriteString(fmt.Sprintf("<w:t>%s</w:t>", escapeXML(text)))
	buf.WriteString("</w:r>")

	buf.WriteString("</w:hyperlink>")
	buf.WriteString("</w:p>")

	return []byte(buf.String())
}

// insertHyperlinkAtPosition inserts hyperlink at the specified position
func (u *Updater) insertHyperlinkAtPosition(docXML, hyperlinkXML []byte, opts HyperlinkOptions) ([]byte, error) {
	switch opts.Position {
	case PositionBeginning:
		return insertAtBodyStart(docXML, hyperlinkXML)
	case PositionEnd:
		return insertAtBodyEnd(docXML, hyperlinkXML)
	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, NewValidationError("anchor", "anchor text required for PositionAfterText")
		}
		return insertAfterText(docXML, hyperlinkXML, opts.Anchor)
	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, NewValidationError("anchor", "anchor text required for PositionBeforeText")
		}
		return insertBeforeText(docXML, hyperlinkXML, opts.Anchor)
	default:
		return insertAtBodyEnd(docXML, hyperlinkXML)
	}
}

// validateURL checks if the URL is valid
func validateURL(urlStr string) error {
	// Check if it's a valid URL
	parsedURL, err := url.Parse(urlStr)
	if err != nil {
		return NewInvalidURLError(urlStr)
	}

	// For web URLs, require a scheme
	if parsedURL.Scheme == "" {
		// Allow mailto: and other common schemes without explicit scheme
		if !strings.HasPrefix(urlStr, "mailto:") &&
			!strings.HasPrefix(urlStr, "file:") &&
			!strings.HasPrefix(urlStr, "ftp:") {
			return NewInvalidURLError(urlStr)
		}
	}

	return nil
}

// escapeXMLAttribute escapes attribute values
func escapeXMLAttribute(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}

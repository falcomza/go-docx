package docxupdater

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// ParagraphStyle defines common paragraph styles
type ParagraphStyle string

const (
	StyleNormal     ParagraphStyle = "Normal"
	StyleHeading1   ParagraphStyle = "Heading1"
	StyleHeading2   ParagraphStyle = "Heading2"
	StyleHeading3   ParagraphStyle = "Heading3"
	StyleTitle      ParagraphStyle = "Title"
	StyleSubtitle   ParagraphStyle = "Subtitle"
	StyleQuote      ParagraphStyle = "Quote"
	StyleIntense    ParagraphStyle = "IntenseQuote"
	StyleListNumber ParagraphStyle = "ListNumber"
	StyleListBullet ParagraphStyle = "ListBullet"
)

// ListType defines the type of list
type ListType string

const (
	ListTypeBullet   ListType = "bullet"   // Bullet list (•)
	ListTypeNumbered ListType = "numbered" // Numbered list (1, 2, 3...)
)

// InsertPosition defines where to insert the paragraph
type InsertPosition int

const (
	// PositionBeginning inserts at the start of the document body
	PositionBeginning InsertPosition = iota
	// PositionEnd inserts at the end of the document body
	PositionEnd
	// PositionAfterText inserts after the first occurrence of specified text
	PositionAfterText
	// PositionBeforeText inserts before the first occurrence of specified text
	PositionBeforeText
)

// ParagraphOptions defines options for paragraph insertion
type ParagraphOptions struct {
	Text      string         // The text content of the paragraph
	Style     ParagraphStyle // The style to apply (default: Normal)
	Position  InsertPosition // Where to insert the paragraph
	Anchor    string         // Text to anchor the insertion (for PositionAfterText/PositionBeforeText)
	Bold      bool           // Make text bold
	Italic    bool           // Make text italic
	Underline bool           // Underline text

	// List properties (alternative to Style-based lists)
	ListType  ListType // Type of list (bullet or numbered)
	ListLevel int      // Indentation level (0-8, default 0)
}

// InsertParagraph inserts a new paragraph into the document
func (u *Updater) InsertParagraph(opts ParagraphOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Text == "" {
		return fmt.Errorf("paragraph text cannot be empty")
	}

	// Default style to Normal if not specified
	if opts.Style == "" {
		opts.Style = StyleNormal
	}

	// Ensure numbering.xml exists if using lists
	if opts.ListType != "" {
		if err := u.ensureNumberingXML(); err != nil {
			return fmt.Errorf("ensure numbering: %w", err)
		}
	}

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Generate paragraph XML
	paraXML := generateParagraphXML(opts)

	// Insert paragraph at the specified position
	updated, err := insertParagraphAtPosition(raw, paraXML, opts)
	if err != nil {
		return fmt.Errorf("insert paragraph: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// InsertParagraphs inserts multiple paragraphs in batch
func (u *Updater) InsertParagraphs(paragraphs []ParagraphOptions) error {
	for i, opts := range paragraphs {
		if err := u.InsertParagraph(opts); err != nil {
			return fmt.Errorf("insert paragraph %d: %w", i, err)
		}
	}
	return nil
}

// generateParagraphXML creates the XML for a paragraph with the specified options
func generateParagraphXML(opts ParagraphOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")

	// Add paragraph properties including style and list numbering
	buf.WriteString("<w:pPr>")

	// Add style if specified
	if opts.Style != StyleNormal {
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, opts.Style))
	}

	// Add numbering properties if ListType is specified
	if opts.ListType != "" {
		var numID int
		if opts.ListType == ListTypeBullet {
			numID = BulletListNumID
		} else if opts.ListType == ListTypeNumbered {
			numID = NumberedListNumID
		}

		if numID > 0 {
			// Validate and constrain list level
			level := min(max(opts.ListLevel, 0), 8)

			buf.WriteString("<w:numPr>")
			buf.WriteString(fmt.Sprintf(`<w:ilvl w:val="%d"/>`, level))
			buf.WriteString(fmt.Sprintf(`<w:numId w:val="%d"/>`, numID))
			buf.WriteString("</w:numPr>")
		}
	}

	buf.WriteString("</w:pPr>")

	// Add text run
	buf.WriteString("<w:r>")

	// Add run properties for formatting
	hasFormatting := opts.Bold || opts.Italic || opts.Underline
	if hasFormatting {
		buf.WriteString("<w:rPr>")
		if opts.Bold {
			buf.WriteString("<w:b/>")
		}
		if opts.Italic {
			buf.WriteString("<w:i/>")
		}
		if opts.Underline {
			buf.WriteString("<w:u w:val=\"single\"/>")
		}
		buf.WriteString("</w:rPr>")
	}

	// Add text content
	buf.WriteString("<w:t")
	// Preserve spaces if text has leading/trailing whitespace
	if strings.HasPrefix(opts.Text, " ") || strings.HasSuffix(opts.Text, " ") {
		buf.WriteString(` xml:space="preserve"`)
	}
	buf.WriteString(">")
	buf.WriteString(xmlEscape(opts.Text))
	buf.WriteString("</w:t>")

	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// insertParagraphAtPosition inserts the paragraph XML at the specified position
func insertParagraphAtPosition(docXML, paraXML []byte, opts ParagraphOptions) ([]byte, error) {
	switch opts.Position {
	case PositionBeginning:
		return insertAtBodyStart(docXML, paraXML)
	case PositionEnd:
		return insertAtBodyEnd(docXML, paraXML)
	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionAfterText")
		}
		return insertAfterText(docXML, paraXML, opts.Anchor)
	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionBeforeText")
		}
		return insertBeforeText(docXML, paraXML, opts.Anchor)
	default:
		return nil, fmt.Errorf("invalid insert position")
	}
}

// insertAtBodyStart inserts paragraph at the start of document body
func insertAtBodyStart(docXML, paraXML []byte) ([]byte, error) {
	// Find <w:body> opening tag
	bodyStart := bytes.Index(docXML, []byte("<w:body>"))
	if bodyStart == -1 {
		return nil, fmt.Errorf("could not find <w:body> tag")
	}

	insertPos := bodyStart + len("<w:body>")

	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:insertPos])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[insertPos:])

	return result, nil
}

// insertAtBodyEnd inserts paragraph at the end of document body (before </w:body>)
func insertAtBodyEnd(docXML, paraXML []byte) ([]byte, error) {
	// Find </w:body> closing tag
	bodyEnd := bytes.Index(docXML, []byte("</w:body>"))
	if bodyEnd == -1 {
		return nil, fmt.Errorf("could not find </w:body> tag")
	}

	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:bodyEnd])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[bodyEnd:])

	return result, nil
}

// insertAfterText inserts paragraph after the paragraph containing the anchor text
func insertAfterText(docXML, paraXML []byte, anchorText string) ([]byte, error) {
	// Find the anchor text in document
	anchorBytes := []byte(xmlEscape(anchorText))
	anchorPos := bytes.Index(docXML, anchorBytes)
	if anchorPos == -1 {
		// Try unescaped
		anchorPos = bytes.Index(docXML, []byte(anchorText))
		if anchorPos == -1 {
			return nil, fmt.Errorf("anchor text %q not found in document", anchorText)
		}
	}

	// Find the end of the paragraph containing this text
	// Search forward for </w:p>
	paraEnd := bytes.Index(docXML[anchorPos:], []byte("</w:p>"))
	if paraEnd == -1 {
		return nil, fmt.Errorf("could not find paragraph end after anchor text")
	}

	insertPos := anchorPos + paraEnd + len("</w:p>")

	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:insertPos])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[insertPos:])

	return result, nil
}

// insertBeforeText inserts paragraph before the paragraph containing the anchor text
func insertBeforeText(docXML, paraXML []byte, anchorText string) ([]byte, error) {
	// Find the anchor text in document
	anchorBytes := []byte(xmlEscape(anchorText))
	anchorPos := bytes.Index(docXML, anchorBytes)
	if anchorPos == -1 {
		// Try unescaped
		anchorPos = bytes.Index(docXML, []byte(anchorText))
		if anchorPos == -1 {
			return nil, fmt.Errorf("anchor text %q not found in document", anchorText)
		}
	}

	// Find the start of the paragraph containing this text
	// Search backward for <w:p>
	searchStart := docXML[:anchorPos]
	paraStart := bytes.LastIndex(searchStart, []byte("<w:p>"))
	if paraStart == -1 {
		// Try with attributes
		paraStart = bytes.LastIndex(searchStart, []byte("<w:p "))
		if paraStart == -1 {
			return nil, fmt.Errorf("could not find paragraph start before anchor text")
		}
	}

	// Insert before this paragraph
	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:paraStart])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[paraStart:])

	return result, nil
}

// AddHeading is a convenience function to add a heading paragraph
func (u *Updater) AddHeading(level int, text string, position InsertPosition) error {
	style := StyleHeading1
	switch level {
	case 1:
		style = StyleHeading1
	case 2:
		style = StyleHeading2
	case 3:
		style = StyleHeading3
	default:
		return fmt.Errorf("heading level must be 1, 2, or 3")
	}

	return u.InsertParagraph(ParagraphOptions{
		Text:     text,
		Style:    style,
		Position: position,
	})
}

// AddText is a convenience function to add normal text paragraph
func (u *Updater) AddText(text string, position InsertPosition) error {
	return u.InsertParagraph(ParagraphOptions{
		Text:     text,
		Style:    StyleNormal,
		Position: position,
	})
}

// AddBulletItem adds a bullet list item at the specified level (0-8)
func (u *Updater) AddBulletItem(text string, level int, position InsertPosition) error {
	return u.InsertParagraph(ParagraphOptions{
		Text:      text,
		ListType:  ListTypeBullet,
		ListLevel: level,
		Position:  position,
	})
}

// AddNumberedItem adds a numbered list item at the specified level (0-8)
func (u *Updater) AddNumberedItem(text string, level int, position InsertPosition) error {
	return u.InsertParagraph(ParagraphOptions{
		Text:      text,
		ListType:  ListTypeNumbered,
		ListLevel: level,
		Position:  position,
	})
}

// AddBulletList adds multiple bullet list items in batch
func (u *Updater) AddBulletList(items []string, level int, position InsertPosition) error {
	paragraphs := make([]ParagraphOptions, len(items))
	for i, item := range items {
		paragraphs[i] = ParagraphOptions{
			Text:      item,
			ListType:  ListTypeBullet,
			ListLevel: level,
			Position:  position,
		}
	}
	return u.InsertParagraphs(paragraphs)
}

// AddNumberedList adds multiple numbered list items in batch
func (u *Updater) AddNumberedList(items []string, level int, position InsertPosition) error {
	paragraphs := make([]ParagraphOptions, len(items))
	for i, item := range items {
		paragraphs[i] = ParagraphOptions{
			Text:      item,
			ListType:  ListTypeNumbered,
			ListLevel: level,
			Position:  position,
		}
	}
	return u.InsertParagraphs(paragraphs)
}

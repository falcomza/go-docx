package docxupdater

import (
	"bytes"
	"fmt"
	"strings"
)

// CaptionType defines the type of caption (Figure or Table)
type CaptionType string

const (
	CaptionFigure CaptionType = "Figure"
	CaptionTable  CaptionType = "Table"
)

// CaptionPosition defines where the caption appears relative to the object
type CaptionPosition string

const (
	CaptionBefore CaptionPosition = "before" // Caption appears before the object
	CaptionAfter  CaptionPosition = "after"  // Caption appears after the object (default for most)
)

// CaptionOptions defines options for caption creation
type CaptionOptions struct {
	// Type of caption (Figure or Table)
	Type CaptionType

	// Position relative to object (before/after)
	Position CaptionPosition

	// Description text after the number (e.g., "Figure 1: <description>")
	Description string

	// Style to use (default: "Caption" which is standard in Word)
	// Can be overridden with custom styles like "Heading 1", "Normal", etc.
	Style string

	// Whether to include automatic numbering using SEQ fields
	// When true, uses Word's built-in field codes for auto-numbering
	AutoNumber bool

	// Alignment options
	Alignment CellAlignment

	// Optional: Manual number override (when AutoNumber is false)
	ManualNumber int
}

// DefaultCaptionOptions returns caption options with sensible defaults
func DefaultCaptionOptions(captionType CaptionType) CaptionOptions {
	position := CaptionAfter
	// Tables commonly have captions above (before) in many styles
	if captionType == CaptionTable {
		position = CaptionBefore
	}

	return CaptionOptions{
		Type:        captionType,
		Position:    position,
		Style:       "Caption",
		AutoNumber:  true,
		Alignment:   CellAlignLeft,
	}
}

// generateCaptionXML creates the XML for a caption paragraph
// MS Word captions use SEQ (sequence) fields for automatic numbering
func generateCaptionXML(opts CaptionOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")

	// Add paragraph properties
	buf.WriteString("<w:pPr>")
	
	// Apply style (default: Caption)
	style := opts.Style
	if style == "" {
		style = "Caption"
	}
	buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, style))

	// Add alignment if specified
	if opts.Alignment != "" {
		alignment := "left"
		switch opts.Alignment {
		case CellAlignCenter:
			alignment = "center"
		case CellAlignRight:
			alignment = "right"
		}
		buf.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, alignment))
	}

	buf.WriteString("</w:pPr>")

	// Add caption label (e.g., "Figure" or "Table") with preserved space
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:t xml:space="preserve">`)
	buf.WriteString(string(opts.Type))
	buf.WriteString(" </w:t>")
	buf.WriteString("</w:r>")

	// Add numbering (either auto SEQ field or manual)
	if opts.AutoNumber {
		// Use SEQ field for automatic numbering
		// This is how Word implements caption numbering
		buf.WriteString(generateSEQFieldXML(opts.Type))
	} else if opts.ManualNumber > 0 {
		// Manual number
		buf.WriteString("<w:r>")
		buf.WriteString("<w:t>")
		buf.WriteString(fmt.Sprintf("%d", opts.ManualNumber))
		buf.WriteString("</w:t>")
		buf.WriteString("</w:r>")
	}

	// Add description if provided
	if opts.Description != "" {
		buf.WriteString("<w:r>")
		buf.WriteString(`<w:t xml:space="preserve">: </w:t>`)
		buf.WriteString("</w:r>")

		buf.WriteString("<w:r>")
		buf.WriteString("<w:t>")
		buf.WriteString(xmlEscape(opts.Description))
		buf.WriteString("</w:t>")
		buf.WriteString("</w:r>")
	}

	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// generateSEQFieldXML creates a SEQ (sequence) field for auto-numbering
// SEQ fields in Word are used for automatic figure/table numbering
// Format: { SEQ Figure \* ARABIC }
func generateSEQFieldXML(captionType CaptionType) string {
	var buf bytes.Buffer

	// Field begin
	buf.WriteString(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`)

	// Field instruction - SEQ field with the caption type identifier
	buf.WriteString(`<w:r><w:instrText xml:space="preserve"> SEQ `)
	buf.WriteString(string(captionType))
	buf.WriteString(` \* ARABIC </w:instrText></w:r>`)

	// Field separator
	buf.WriteString(`<w:r><w:fldChar w:fldCharType="separate"/></w:r>`)

	// Field result (placeholder - Word will calculate this)
	// Default to "1" as placeholder; Word will recalculate on open
	buf.WriteString(`<w:r><w:t>1</w:t></w:r>`)

	// Field end
	buf.WriteString(`<w:r><w:fldChar w:fldCharType="end"/></w:r>`)

	return buf.String()
}

// insertCaptionWithElement inserts a caption along with an element (chart/table)
// The caption can be before or after the element based on CaptionPosition
func insertCaptionWithElement(docXML, captionXML, elementXML []byte, position CaptionPosition) []byte {
	var buf bytes.Buffer

	if position == CaptionBefore {
		// Caption first, then element
		buf.Write(captionXML)
		buf.Write(elementXML)
	} else {
		// Element first, then caption
		buf.Write(elementXML)
		buf.Write(captionXML)
	}

	return buf.Bytes()
}

// ValidateCaptionOptions validates caption options
func ValidateCaptionOptions(opts *CaptionOptions) error {
	if opts == nil {
		return nil // No caption requested
	}

	// Validate caption type
	if opts.Type != CaptionFigure && opts.Type != CaptionTable {
		return fmt.Errorf("invalid caption type: %s (must be 'Figure' or 'Table')", opts.Type)
	}

	// Validate position
	if opts.Position != "" && opts.Position != CaptionBefore && opts.Position != CaptionAfter {
		return fmt.Errorf("invalid caption position: %s (must be 'before' or 'after')", opts.Position)
	}

	// Set default position if not specified
	if opts.Position == "" {
		if opts.Type == CaptionTable {
			opts.Position = CaptionBefore
		} else {
			opts.Position = CaptionAfter
		}
	}

	// Description is optional but should be reasonable length if present
	if len(opts.Description) > 500 {
		return fmt.Errorf("caption description too long: %d characters (max 500)", len(opts.Description))
	}

	return nil
}

// FormatCaptionText formats a caption text for display (without XML)
// Useful for previewing what the caption will look like
func FormatCaptionText(opts CaptionOptions) string {
	var buf strings.Builder

	buf.WriteString(string(opts.Type))
	buf.WriteString(" ")

	if opts.AutoNumber {
		buf.WriteString("#") // Placeholder for auto number
	} else if opts.ManualNumber > 0 {
		buf.WriteString(fmt.Sprintf("%d", opts.ManualNumber))
	}

	if opts.Description != "" {
		buf.WriteString(": ")
		buf.WriteString(opts.Description)
	}

	return buf.String()
}

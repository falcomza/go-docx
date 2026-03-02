package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// StyleType defines the type of style
type StyleType string

const (
	// StyleTypeParagraph is a paragraph style (affects entire paragraph)
	StyleTypeParagraph StyleType = "paragraph"
	// StyleTypeCharacter is a character style (affects inline text runs)
	StyleTypeCharacter StyleType = "character"
)

// StyleDefinition defines a custom style to add to the document
type StyleDefinition struct {
	// ID is the style ID used for referencing (e.g., "CustomHeading")
	ID string

	// Name is the display name shown in Word's style gallery
	Name string

	// Type is the style type (paragraph or character)
	Type StyleType

	// BasedOn is the parent style ID (e.g., "Normal", "Heading1")
	BasedOn string

	// NextStyle is the style applied to the next paragraph (paragraph styles only)
	NextStyle string

	// Font settings
	FontFamily string // e.g., "Arial", "Times New Roman"
	FontSize   int    // Font size in half-points (e.g., 24 = 12pt)
	Color      string // Hex color code without '#' (e.g., "FF0000")

	// Text formatting
	Bold          bool
	Italic        bool
	Underline     bool
	Strikethrough bool
	AllCaps       bool
	SmallCaps     bool

	// Paragraph formatting (paragraph styles only)
	Alignment    ParagraphAlignment
	SpaceBefore  int // Space before paragraph in twips
	SpaceAfter   int // Space after paragraph in twips
	LineSpacing  int // Line spacing in 240ths of a line (e.g., 240 = single, 480 = double)
	IndentLeft   int // Left indent in twips
	IndentRight  int // Right indent in twips
	IndentFirst  int // First line indent in twips
	KeepNext     bool
	KeepLines    bool
	PageBreakBef bool

	// Outline level (0-8, paragraph styles only, used for TOC)
	OutlineLevel int
}

// AddStyle adds a custom style definition to the document.
// The style can then be used with InsertParagraph by setting
// ParagraphOptions.Style to the style ID.
func (u *Updater) AddStyle(def StyleDefinition) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if def.ID == "" {
		return fmt.Errorf("style ID cannot be empty")
	}
	if def.Name == "" {
		def.Name = def.ID
	}
	if def.Type == "" {
		def.Type = StyleTypeParagraph
	}

	styleXML := generateStyleXML(def)

	stylesPath := filepath.Join(u.tempDir, "word", "styles.xml")

	// Check if styles.xml exists
	raw, err := os.ReadFile(stylesPath)
	if err != nil {
		// Create new styles.xml
		updated := generateStylesDocument(styleXML)
		if err := atomicWriteFile(stylesPath, updated, 0o644); err != nil {
			return fmt.Errorf("write styles.xml: %w", err)
		}
		// Ensure relationship and content type
		if err := u.ensureStylesRelationship(); err != nil {
			return fmt.Errorf("ensure styles relationship: %w", err)
		}
		return nil
	}

	// Inject style into existing styles.xml
	updated, err := injectStyle(raw, styleXML)
	if err != nil {
		return fmt.Errorf("inject style: %w", err)
	}

	if err := atomicWriteFile(stylesPath, updated, 0o644); err != nil {
		return fmt.Errorf("write styles.xml: %w", err)
	}

	return nil
}

// AddStyles adds multiple custom style definitions in batch.
func (u *Updater) AddStyles(defs []StyleDefinition) error {
	for i, def := range defs {
		if err := u.AddStyle(def); err != nil {
			return fmt.Errorf("add style %d (%s): %w", i, def.ID, err)
		}
	}
	return nil
}

// generateStyleXML creates the XML for a single style definition
func generateStyleXML(def StyleDefinition) []byte {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`<w:style w:type="%s" w:styleId="%s">`,
		xmlEscape(string(def.Type)), xmlEscape(def.ID)))

	buf.WriteString(fmt.Sprintf(`<w:name w:val="%s"/>`, xmlEscape(def.Name)))

	if def.BasedOn != "" {
		buf.WriteString(fmt.Sprintf(`<w:basedOn w:val="%s"/>`, xmlEscape(def.BasedOn)))
	}
	if def.NextStyle != "" {
		buf.WriteString(fmt.Sprintf(`<w:next w:val="%s"/>`, xmlEscape(def.NextStyle)))
	}

	// Paragraph properties (only for paragraph styles)
	if def.Type == StyleTypeParagraph {
		pPr := generateStyleParagraphProps(def)
		if pPr != "" {
			buf.WriteString(pPr)
		}
	}

	// Run properties (for both paragraph and character styles)
	rPr := generateStyleRunProps(def)
	if rPr != "" {
		buf.WriteString(rPr)
	}

	buf.WriteString("</w:style>")

	return buf.Bytes()
}

// generateStyleParagraphProps creates paragraph properties XML for a style
func generateStyleParagraphProps(def StyleDefinition) string {
	var buf strings.Builder
	hasProps := false

	var inner strings.Builder

	if alignment, ok := paragraphAlignmentValue(def.Alignment); ok {
		inner.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, alignment))
		hasProps = true
	}

	if def.SpaceBefore > 0 || def.SpaceAfter > 0 || def.LineSpacing > 0 {
		inner.WriteString("<w:spacing")
		if def.SpaceBefore > 0 {
			inner.WriteString(fmt.Sprintf(` w:before="%d"`, def.SpaceBefore))
		}
		if def.SpaceAfter > 0 {
			inner.WriteString(fmt.Sprintf(` w:after="%d"`, def.SpaceAfter))
		}
		if def.LineSpacing > 0 {
			inner.WriteString(fmt.Sprintf(` w:line="%d" w:lineRule="auto"`, def.LineSpacing))
		}
		inner.WriteString("/>")
		hasProps = true
	}

	if def.IndentLeft > 0 || def.IndentRight > 0 || def.IndentFirst != 0 {
		inner.WriteString("<w:ind")
		if def.IndentLeft > 0 {
			inner.WriteString(fmt.Sprintf(` w:left="%d"`, def.IndentLeft))
		}
		if def.IndentRight > 0 {
			inner.WriteString(fmt.Sprintf(` w:right="%d"`, def.IndentRight))
		}
		if def.IndentFirst != 0 {
			inner.WriteString(fmt.Sprintf(` w:firstLine="%d"`, def.IndentFirst))
		}
		inner.WriteString("/>")
		hasProps = true
	}

	if def.KeepNext {
		inner.WriteString("<w:keepNext/>")
		hasProps = true
	}
	if def.KeepLines {
		inner.WriteString("<w:keepLines/>")
		hasProps = true
	}
	if def.PageBreakBef {
		inner.WriteString("<w:pageBreakBefore/>")
		hasProps = true
	}

	if def.OutlineLevel > 0 && def.OutlineLevel <= 9 {
		inner.WriteString(fmt.Sprintf(`<w:outlineLvl w:val="%d"/>`, def.OutlineLevel-1))
		hasProps = true
	}

	if hasProps {
		buf.WriteString("<w:pPr>")
		buf.WriteString(inner.String())
		buf.WriteString("</w:pPr>")
	}

	return buf.String()
}

// generateStyleRunProps creates run properties XML for a style
func generateStyleRunProps(def StyleDefinition) string {
	var buf strings.Builder
	hasProps := false

	var inner strings.Builder

	if def.FontFamily != "" {
		escaped := xmlEscape(def.FontFamily)
		inner.WriteString(fmt.Sprintf(`<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>`,
			escaped, escaped, escaped))
		hasProps = true
	}

	if def.Bold {
		inner.WriteString("<w:b/>")
		hasProps = true
	}
	if def.Italic {
		inner.WriteString("<w:i/>")
		hasProps = true
	}
	if def.Underline {
		inner.WriteString(`<w:u w:val="single"/>`)
		hasProps = true
	}
	if def.Strikethrough {
		inner.WriteString("<w:strike/>")
		hasProps = true
	}
	if def.AllCaps {
		inner.WriteString("<w:caps/>")
		hasProps = true
	}
	if def.SmallCaps {
		inner.WriteString("<w:smallCaps/>")
		hasProps = true
	}

	if def.FontSize > 0 {
		inner.WriteString(fmt.Sprintf(`<w:sz w:val="%d"/>`, def.FontSize))
		inner.WriteString(fmt.Sprintf(`<w:szCs w:val="%d"/>`, def.FontSize))
		hasProps = true
	}

	if def.Color != "" {
		color := normalizeHexColor(def.Color)
		if color == "" {
			color = def.Color
		}
		inner.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, xmlEscape(color)))
		hasProps = true
	}

	if hasProps {
		buf.WriteString("<w:rPr>")
		buf.WriteString(inner.String())
		buf.WriteString("</w:rPr>")
	}

	return buf.String()
}

// generateStylesDocument creates a complete styles.xml with a single style
func generateStylesDocument(styleXML []byte) []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)
	buf.WriteString("\n")
	buf.Write(styleXML)
	buf.WriteString("\n")
	buf.WriteString(`</w:styles>`)

	return buf.Bytes()
}

// injectStyle inserts a style definition into an existing styles.xml
func injectStyle(stylesXML, styleXML []byte) ([]byte, error) {
	closeTag := []byte("</w:styles>")
	closeIdx := bytes.LastIndex(stylesXML, closeTag)
	if closeIdx == -1 {
		return nil, fmt.Errorf("could not find </w:styles> closing tag")
	}

	result := make([]byte, 0, len(stylesXML)+len(styleXML)+1)
	result = append(result, stylesXML[:closeIdx]...)
	result = append(result, styleXML...)
	result = append(result, '\n')
	result = append(result, stylesXML[closeIdx:]...)

	return result, nil
}

// ensureStylesRelationship ensures the styles.xml relationship and content type exist
func (u *Updater) ensureStylesRelationship() error {
	// Check/add relationship
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read rels: %w", err)
	}

	content := string(raw)
	if !strings.Contains(content, "styles.xml") {
		relID, err := getNextRelIDFromFile(relsPath)
		if err != nil {
			return fmt.Errorf("get next rel ID: %w", err)
		}
		newRel := fmt.Sprintf(
			`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
			relID,
		)
		content = strings.Replace(content, "</Relationships>", newRel+"</Relationships>", 1)
		if err := atomicWriteFile(relsPath, []byte(content), 0o644); err != nil {
			return fmt.Errorf("write rels: %w", err)
		}
	}

	// Check/add content type
	ctPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	ctRaw, err := os.ReadFile(ctPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	ctContent := string(ctRaw)
	if !strings.Contains(ctContent, "styles.xml") {
		override := `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`
		ctContent = strings.Replace(ctContent, "</Types>", override+"</Types>", 1)
		if err := atomicWriteFile(ctPath, []byte(ctContent), 0o644); err != nil {
			return fmt.Errorf("write content types: %w", err)
		}
	}

	return nil
}

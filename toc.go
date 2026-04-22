package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
)

// TOCOptions defines options for Table of Contents
type TOCOptions struct {
	// Title for the TOC (default: "Table of Contents")
	Title string

	// Outline levels to include (default: "1-3")
	// Example: "1-3" includes Heading1, Heading2, Heading3
	OutlineLevels string

	// Position where to insert the TOC
	Position InsertPosition

	// Anchor text for position-based insertion
	Anchor string
}

// CaptionListOptions defines options for Table of Figures / Table of Tables.
type CaptionListOptions struct {
	// Title for the generated list (e.g. "Table of Figures"). Empty means no title paragraph.
	Title string

	// Position where to insert the field.
	Position InsertPosition

	// Anchor text for position-based insertion.
	Anchor string
}

// DefaultTOCOptions returns default TOC options
func DefaultTOCOptions() TOCOptions {
	return TOCOptions{
		Title:         "Table of Contents",
		OutlineLevels: "1-3",
		Position:      PositionBeginning,
	}
}

// DefaultTableOfFiguresOptions returns default options for a Table of Figures.
func DefaultTableOfFiguresOptions() CaptionListOptions {
	return CaptionListOptions{
		Title:    "Table of Figures",
		Position: PositionBeginning,
	}
}

// DefaultTableOfTablesOptions returns default options for a Table of Tables.
func DefaultTableOfTablesOptions() CaptionListOptions {
	return CaptionListOptions{
		Title:    "Table of Tables",
		Position: PositionBeginning,
	}
}

// InsertTOC inserts a Table of Contents field into the document.
// The TOC uses Word field codes and will be populated when the document
// is opened in Word and the user updates the field (Ctrl+A, F9).
func (u *Updater) InsertTOC(opts TOCOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	if opts.OutlineLevels == "" {
		opts.OutlineLevels = "1-3"
	}

	tocXML := generateTOCXML(opts)

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := insertTOCAtPosition(raw, tocXML, opts)
	if err != nil {
		return fmt.Errorf("insert TOC: %w", err)
	}

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// InsertTableOfFigures inserts a caption-based list field for Figure captions.
// Word populates the entries when the field is updated.
func (u *Updater) InsertTableOfFigures(opts CaptionListOptions) error {
	return u.insertCaptionList(opts, CaptionFigure)
}

// InsertTableOfTables inserts a caption-based list field for Table captions.
// Word populates the entries when the field is updated.
func (u *Updater) InsertTableOfTables(opts CaptionListOptions) error {
	return u.insertCaptionList(opts, CaptionTable)
}

func (u *Updater) insertCaptionList(opts CaptionListOptions, captionType CaptionType) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	listXML := generateCaptionListXML(opts, captionType)

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := insertTOCAtPosition(raw, listXML, TOCOptions{
		Position: opts.Position,
		Anchor:   opts.Anchor,
	})
	if err != nil {
		return fmt.Errorf("insert caption list: %w", err)
	}

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// generateTOCXML creates the XML for a Table of Contents field.
// The output includes an optional title paragraph followed by the TOC field paragraph.
func generateTOCXML(opts TOCOptions) []byte {
	var buf bytes.Buffer

	// Title paragraph comes BEFORE the TOC field
	if opts.Title != "" {
		buf.Write(generateTOCTitleXML(opts.Title))
	}

	// Build the TOC field instruction.
	// Word field switches use single backslashes:
	//   \o "1-3" - include outline levels 1-3
	//   \h       - make entries hyperlinks
	//   \z       - hide tab leaders in Web Layout view
	//   \u       - use applied paragraph outline level
	fieldInstr := fmt.Sprintf(` TOC \o "%s" \h \z \u `, opts.OutlineLevels)

	// TOC field paragraph
	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr/>")

	// Field begin
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:fldChar w:fldCharType="begin"/>`)
	buf.WriteString("</w:r>")

	// Field instruction
	buf.WriteString("<w:r>")
	buf.WriteString(fmt.Sprintf(`<w:instrText xml:space="preserve">%s</w:instrText>`, xmlEscape(fieldInstr)))
	buf.WriteString("</w:r>")

	// Field separate (marks end of instruction, start of result)
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:fldChar w:fldCharType="separate"/>`)
	buf.WriteString("</w:r>")

	// Placeholder result text
	buf.WriteString("<w:r>")
	buf.WriteString("<w:rPr><w:i/></w:rPr>")
	buf.WriteString("<w:t>Update this field to show Table of Contents</w:t>")
	buf.WriteString("</w:r>")

	// Field end
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:fldChar w:fldCharType="end"/>`)
	buf.WriteString("</w:r>")

	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// generateCaptionListXML creates the XML for a caption-based list field such as
// a Table of Figures or Table of Tables.
func generateCaptionListXML(opts CaptionListOptions, captionType CaptionType) []byte {
	var buf bytes.Buffer

	if opts.Title != "" {
		buf.Write(generateTOCTitleXML(opts.Title))
	}

	listName := "Table of Figures"
	if captionType == CaptionTable {
		listName = "Table of Tables"
	}
	fieldInstr := fmt.Sprintf(` TOC \h \z \c "%s" `, captionType)

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr/>")
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:fldChar w:fldCharType="begin"/>`)
	buf.WriteString("</w:r>")
	buf.WriteString("<w:r>")
	buf.WriteString(fmt.Sprintf(`<w:instrText xml:space="preserve">%s</w:instrText>`, xmlEscape(fieldInstr)))
	buf.WriteString("</w:r>")
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:fldChar w:fldCharType="separate"/>`)
	buf.WriteString("</w:r>")
	buf.WriteString("<w:r>")
	buf.WriteString("<w:rPr><w:i/></w:rPr>")
	buf.WriteString(fmt.Sprintf(`<w:t>Update this field to show %s</w:t>`, xmlEscape(listName)))
	buf.WriteString("</w:r>")
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:fldChar w:fldCharType="end"/>`)
	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// generateTOCTitleXML creates the XML for the TOC title paragraph
func generateTOCTitleXML(title string) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr>")
	buf.WriteString(`<w:pStyle w:val="TOCHeading"/>`)
	buf.WriteString(`<w:jc w:val="center"/>`)
	buf.WriteString("</w:pPr>")
	buf.WriteString("<w:r>")
	buf.WriteString("<w:rPr>")
	buf.WriteString("<w:b/>")
	buf.WriteString(`<w:sz w:val="44"/>`)
	buf.WriteString("</w:rPr>")
	buf.WriteString(fmt.Sprintf(`<w:t xml:space="preserve">%s</w:t>`, xmlEscape(title)))
	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// insertTOCAtPosition inserts the TOC XML at the specified position
func insertTOCAtPosition(docXML, tocXML []byte, opts TOCOptions) ([]byte, error) {
	var insertPos int
	var err error

	switch opts.Position {
	case PositionBeginning:
		insertPos, err = findBodyContentStart(docXML)
		if err != nil {
			return nil, err
		}
	case PositionEnd:
		bodyEnd := bytes.Index(docXML, []byte("</w:body>"))
		if bodyEnd == -1 {
			return nil, fmt.Errorf("could not find </w:body> tag")
		}
		if sectPrPos := bytes.LastIndex(docXML[:bodyEnd], []byte("<w:sectPr")); sectPrPos != -1 {
			insertPos = sectPrPos
		} else {
			insertPos = bodyEnd
		}
	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionAfterText")
		}
		_, insertPos, err = findParagraphRangeByAnchor(docXML, opts.Anchor)
		if err != nil {
			return nil, err
		}
	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionBeforeText")
		}
		insertPos, _, err = findParagraphRangeByAnchor(docXML, opts.Anchor)
		if err != nil {
			return nil, err
		}
	default:
		return nil, fmt.Errorf("invalid insert position")
	}

	result := make([]byte, 0, len(docXML)+len(tocXML))
	result = append(result, docXML[:insertPos]...)
	result = append(result, tocXML...)
	result = append(result, docXML[insertPos:]...)

	return result, nil
}

// UpdateTOC marks an existing Table of Contents for update.
// When the document is opened in Word, it will prompt the user to
// update the TOC field to reflect the current document headings.
func (u *Updater) UpdateTOC() error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated := markTOCForUpdate(raw)

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// markTOCForUpdate finds the TOC field's begin fldChar and adds the
// w:dirty="true" attribute, which tells Word to recalculate the TOC
// when the document is opened.
func markTOCForUpdate(docXML []byte) []byte {
	// Find instrText containing a TOC field code
	tocIdx := bytes.Index(docXML, []byte(">TOC"))
	if tocIdx == -1 {
		tocIdx = bytes.Index(docXML, []byte("> TOC"))
	}
	if tocIdx == -1 {
		return docXML
	}

	// Find the nearest fldChar begin before the TOC instrText
	searchArea := docXML[:tocIdx]
	oldBegin := []byte(`w:fldCharType="begin"/>`)
	beginIdx := bytes.LastIndex(searchArea, oldBegin)
	if beginIdx == -1 {
		return docXML
	}

	// Check if dirty attribute already present
	region := docXML[beginIdx:tocIdx]
	if bytes.Contains(region, []byte("w:dirty")) {
		return docXML
	}

	// Replace begin fldChar with dirty version
	newBegin := []byte(`w:fldCharType="begin" w:dirty="true"/>`)

	result := make([]byte, 0, len(docXML)+20)
	result = append(result, docXML[:beginIdx]...)
	result = append(result, newBegin...)
	result = append(result, docXML[beginIdx+len(oldBegin):]...)

	return result
}

// GetTOCEntries extracts TOC entries from the document.
// Returns a slice of TOCEntry with level and text.
func (u *Updater) GetTOCEntries() ([]TOCEntry, error) {
	if u == nil {
		return nil, fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return nil, fmt.Errorf("read document.xml: %w", err)
	}

	return parseTOCEntries(raw), nil
}

// TOCEntry represents an entry in the Table of Contents
type TOCEntry struct {
	Level int    // Heading level (1-9)
	Text  string // Entry text
	Page  int    // Page number (if available)
}

// parseTOCEntries extracts TOC entries from document XML by looking
// for paragraphs with TOC styles (TOC1, TOC2, TOC3, etc.)
func parseTOCEntries(docXML []byte) []TOCEntry {
	var entries []TOCEntry

	searchStart := 0
	for {
		paraStart := bytes.Index(docXML[searchStart:], []byte("<w:p"))
		if paraStart == -1 {
			break
		}
		paraStart += searchStart

		paraEnd := bytes.Index(docXML[paraStart:], []byte("</w:p>"))
		if paraEnd == -1 {
			break
		}
		paraEnd += paraStart + len("</w:p>")

		paraXML := docXML[paraStart:paraEnd]

		// Check for TOC paragraph styles
		level := -1
		if bytes.Contains(paraXML, []byte(`w:val="TOC1"`)) {
			level = 1
		} else if bytes.Contains(paraXML, []byte(`w:val="TOC2"`)) {
			level = 2
		} else if bytes.Contains(paraXML, []byte(`w:val="TOC3"`)) {
			level = 3
		} else if bytes.Contains(paraXML, []byte(`w:val="TOC4"`)) {
			level = 4
		} else if bytes.Contains(paraXML, []byte(`w:val="TOC5"`)) {
			level = 5
		}

		if level > 0 {
			text := extractParagraphPlainText(paraXML)
			if text != "" {
				entries = append(entries, TOCEntry{
					Level: level,
					Text:  text,
				})
			}
		}

		searchStart = paraEnd
	}

	return entries
}

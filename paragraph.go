package godocx

import (
	"bytes"
	"fmt"
	"html"
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

// ParagraphAlignment defines paragraph text alignment.
type ParagraphAlignment string

const (
	ParagraphAlignLeft    ParagraphAlignment = "left"
	ParagraphAlignCenter  ParagraphAlignment = "center"
	ParagraphAlignRight   ParagraphAlignment = "right"
	ParagraphAlignJustify ParagraphAlignment = "both"
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

// RunOptions defines formatting and content for a single text run within a paragraph.
// A run is the smallest unit of text in OpenXML that can carry its own character formatting.
// Use multiple RunOptions in ParagraphOptions.Runs to mix bold, italic, colored, and
// differently-sized text within the same paragraph.
type RunOptions struct {
	// Text is the content of this run. Newlines (\n) become <w:br/> and tabs (\t) become <w:tab/>.
	Text string

	// Character formatting
	Bold          bool
	Italic        bool
	Underline     bool // Single underline
	Strikethrough bool // Strikethrough text
	Superscript   bool // Raise text above baseline
	Subscript     bool // Lower text below baseline

	// Color is a 6-digit hex RGB value without '#', e.g. "FF0000" for red.
	Color string

	// Highlight is a named highlight color: "yellow", "green", "cyan", "magenta",
	// "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta",
	// "darkRed", "darkYellow", "darkGray", "lightGray", "black".
	Highlight string

	// FontSize is the font size in points (e.g. 12.0). Zero means inherit.
	FontSize float64

	// FontName sets the ASCII/Unicode font (e.g. "Arial", "Times New Roman").
	FontName string

	// URL sets an inline hyperlink on this run. When non-empty the run is emitted
	// as a <w:hyperlink> element. Default hyperlink styling (blue, underlined) is
	// applied unless Color or Underline are explicitly set on the run.
	URL string
}

// ParagraphOptions defines options for paragraph insertion
type ParagraphOptions struct {
	// Text is used when Runs is empty — the whole paragraph gets a single run with
	// the Bold/Italic/Underline flags below applied uniformly.
	Text      string         // The text content of the paragraph
	Style     ParagraphStyle // The style to apply (default: Normal)
	Alignment ParagraphAlignment
	Position  InsertPosition // Where to insert the paragraph
	Anchor    string         // Text to anchor the insertion (for PositionAfterText/PositionBeforeText)

	// Single-run formatting flags — only used when Runs is empty.
	Bold      bool // Make text bold
	Italic    bool // Make text italic
	Underline bool // Underline text

	// Runs allows building a multi-run paragraph where each run can have independent
	// character formatting. When Runs is non-empty, Text/Bold/Italic/Underline above
	// are ignored.
	Runs []RunOptions

	// List properties (alternative to Style-based lists)
	ListType  ListType // Type of list (bullet or numbered)
	ListLevel int      // Indentation level (0-8, default 0)

	// Pagination control
	KeepNext  bool // Keep this paragraph on the same page as the next (prevents orphaned headings)
	KeepLines bool // Keep all lines of this paragraph together on the same page
}

type listNumberingIDs struct {
	bulletNumID   int
	numberedNumID int
}

// InsertParagraph inserts a new paragraph into the document
func (u *Updater) InsertParagraph(opts ParagraphOptions) error {
	if u == nil {
		return &DocxError{Code: ErrCodeValidation, Message: "updater is nil"}
	}
	if opts.Text == "" && len(opts.Runs) == 0 {
		return NewValidationError("text", "paragraph text cannot be empty: provide Text or at least one Run")
	}

	// Default style to Normal if not specified
	if opts.Style == "" {
		opts.Style = StyleNormal
	}

	listIDs := listNumberingIDs{bulletNumID: BulletListNumID, numberedNumID: NumberedListNumID}

	// Ensure numbering.xml exists if using lists
	if opts.ListType != "" {
		if err := u.ensureNumberingXML(); err != nil {
			return fmt.Errorf("ensure numbering: %w", err)
		}
		listIDs = u.getListNumberingIDs()
	}

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Pre-resolve URL relationships for any inline hyperlink runs.
	urlRelIDs := make(map[string]string)
	for _, run := range opts.Runs {
		if run.URL != "" {
			if _, seen := urlRelIDs[run.URL]; !seen {
				rID, err := u.addHyperlinkRelationship(run.URL)
				if err != nil {
					return fmt.Errorf("register hyperlink for %q: %w", run.URL, err)
				}
				urlRelIDs[run.URL] = rID
			}
		}
	}

	// Generate paragraph XML
	paraXML := generateParagraphXML(opts, listIDs, urlRelIDs)

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

// InsertParagraphs inserts multiple paragraphs in a single read-modify-write pass,
// which is significantly more efficient than calling InsertParagraph N times.
func (u *Updater) InsertParagraphs(paragraphs []ParagraphOptions) error {
	if u == nil {
		return &DocxError{Code: ErrCodeValidation, Message: "updater is nil"}
	}
	if len(paragraphs) == 0 {
		return nil
	}

	// Validate all paragraphs upfront before touching any files.
	for i, opts := range paragraphs {
		if opts.Text == "" && len(opts.Runs) == 0 {
			return fmt.Errorf("paragraph %d: %w", i,
				NewValidationError("text", "paragraph text cannot be empty: provide Text or at least one Run"))
		}
	}

	// Ensure numbering.xml exists once if any paragraph uses a list.
	for _, opts := range paragraphs {
		if opts.ListType != "" {
			if err := u.ensureNumberingXML(); err != nil {
				return fmt.Errorf("ensure numbering: %w", err)
			}
			break
		}
	}
	listIDs := u.getListNumberingIDs()

	// Pre-register all URL relationships in a single pass.
	urlRelIDs := make(map[string]string)
	for _, opts := range paragraphs {
		for _, run := range opts.Runs {
			if run.URL != "" {
				if _, seen := urlRelIDs[run.URL]; !seen {
					rID, err := u.addHyperlinkRelationship(run.URL)
					if err != nil {
						return fmt.Errorf("register hyperlink for %q: %w", run.URL, err)
					}
					urlRelIDs[run.URL] = rID
				}
			}
		}
	}

	// Read document.xml once.
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Apply all insertions in memory.
	for i, opts := range paragraphs {
		if opts.Style == "" {
			opts.Style = StyleNormal
		}
		paraXML := generateParagraphXML(opts, listIDs, urlRelIDs)
		raw, err = insertParagraphAtPosition(raw, paraXML, opts)
		if err != nil {
			return fmt.Errorf("insert paragraph %d: %w", i, err)
		}
	}

	// Write document.xml once.
	if err := os.WriteFile(docPath, raw, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}
	return nil
}

// generateParagraphXML creates the XML for a paragraph with the specified options.
// urlRelIDs maps URL strings to their relationship IDs (returned by addHyperlinkRelationship).
// Runs with a non-empty URL are emitted as inline <w:hyperlink> elements when a
// corresponding relationship ID exists in urlRelIDs; otherwise they fall back to plain runs.
func generateParagraphXML(opts ParagraphOptions, listIDs listNumberingIDs, urlRelIDs map[string]string) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")

	// Add paragraph properties including style and list numbering
	buf.WriteString("<w:pPr>")

	// Add style if specified
	if opts.Style != StyleNormal {
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, xmlEscape(string(opts.Style))))
	}

	if alignment, ok := paragraphAlignmentValue(opts.Alignment); ok {
		buf.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, alignment))
	}

	// Add numbering properties if ListType is specified
	if opts.ListType != "" {
		var numID int
		if opts.ListType == ListTypeBullet {
			numID = listIDs.bulletNumID
		} else if opts.ListType == ListTypeNumbered {
			numID = listIDs.numberedNumID
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

	// Pagination control: keep with next paragraph (headings) and keep lines together.
	if opts.KeepNext {
		buf.WriteString("<w:keepNext/>")
	}
	if opts.KeepLines {
		buf.WriteString("<w:keepLines/>")
	}

	buf.WriteString("</w:pPr>")

	if len(opts.Runs) > 0 {
		// Multi-run paragraph: emit one <w:r> per RunOptions entry.
		// Runs with a URL are wrapped in <w:hyperlink> when a relationship ID is available.
		for _, run := range opts.Runs {
			if run.URL != "" {
				if rID, ok := urlRelIDs[run.URL]; ok {
					writeHyperlinkRunXML(&buf, run, rID)
					continue
				}
			}
			writeRunXML(&buf, run)
		}
	} else {
		// Legacy single-run paragraph using top-level Text/Bold/Italic/Underline.
		legacyRun := RunOptions{
			Text:      opts.Text,
			Bold:      opts.Bold,
			Italic:    opts.Italic,
			Underline: opts.Underline,
		}
		writeRunXML(&buf, legacyRun)
	}

	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// writeRunXML emits a full <w:r>...</w:r> element for the given RunOptions.
func writeRunXML(buf *bytes.Buffer, run RunOptions) {
	buf.WriteString("<w:r>")

	hasRPr := run.Bold || run.Italic || run.Underline || run.Strikethrough ||
		run.Superscript || run.Subscript ||
		run.Color != "" || run.Highlight != "" ||
		run.FontSize > 0 || run.FontName != ""

	if hasRPr {
		buf.WriteString("<w:rPr>")
		if run.FontName != "" {
			buf.WriteString(fmt.Sprintf(
				`<w:rFonts w:ascii="%s" w:hAnsi="%s"/>`,
				xmlEscape(run.FontName), xmlEscape(run.FontName),
			))
		}
		if run.Bold {
			buf.WriteString("<w:b/>")
		}
		if run.Italic {
			buf.WriteString("<w:i/>")
		}
		if run.Strikethrough {
			buf.WriteString("<w:strike/>")
		}
		if run.Color != "" {
			// normalizeHexColor validates and normalises the value; skip invalid strings
			// to avoid emitting malformed XML attribute values.
			if normalized := normalizeHexColor(run.Color); normalized != "" {
				buf.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, normalized))
			}
		}
		if run.Highlight != "" {
			buf.WriteString(fmt.Sprintf(`<w:highlight w:val="%s"/>`, xmlEscape(run.Highlight)))
		}
		if run.Underline {
			buf.WriteString(`<w:u w:val="single"/>`)
		}
		if run.FontSize > 0 {
			// w:sz / w:szCs values are in half-points (see FontSizeHalfPointsFactor).
			hp := int(run.FontSize * FontSizeHalfPointsFactor)
			buf.WriteString(fmt.Sprintf(`<w:sz w:val="%d"/>`, hp))
			buf.WriteString(fmt.Sprintf(`<w:szCs w:val="%d"/>`, hp))
		}
		if run.Superscript {
			buf.WriteString(`<w:vertAlign w:val="superscript"/>`)
		} else if run.Subscript {
			buf.WriteString(`<w:vertAlign w:val="subscript"/>`)
		}
		buf.WriteString("</w:rPr>")
	}

	writeRunTextWithControls(buf, run.Text)

	buf.WriteString("</w:r>")
}

// writeHyperlinkRunXML emits a <w:hyperlink> element wrapping a styled run.
// Default hyperlink styling (blue, underlined) is applied unless the run overrides them.
func writeHyperlinkRunXML(buf *bytes.Buffer, run RunOptions, rID string) {
	buf.WriteString(fmt.Sprintf(`<w:hyperlink r:id="%s" w:history="1">`, xmlEscape(rID)))
	if run.Color == "" {
		run.Color = "0563C1"
	}
	if !run.Underline {
		run.Underline = true
	}
	writeRunXML(buf, run)
	buf.WriteString("</w:hyperlink>")
}

func writeRunTextWithControls(buf *bytes.Buffer, text string) {
	start := 0
	flushText := func(seg string) {
		if seg == "" {
			return
		}
		buf.WriteString("<w:t")
		if strings.HasPrefix(seg, " ") || strings.HasSuffix(seg, " ") {
			buf.WriteString(` xml:space="preserve"`)
		}
		buf.WriteString(">")
		buf.WriteString(xmlEscape(seg))
		buf.WriteString("</w:t>")
	}

	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '\n':
			flushText(text[start:i])
			buf.WriteString("<w:br/>")
			start = i + 1
		case '\t':
			flushText(text[start:i])
			buf.WriteString("<w:tab/>")
			start = i + 1
		}
	}
	flushText(text[start:])
}

func paragraphAlignmentValue(alignment ParagraphAlignment) (string, bool) {
	switch strings.ToLower(string(alignment)) {
	case "left":
		return "left", true
	case "center":
		return "center", true
	case "right":
		return "right", true
	case "both", "justify":
		return "both", true
	default:
		return "", false
	}
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
	bodyContentStart, err := findBodyContentStart(docXML)
	if err != nil {
		return nil, err
	}

	insertPos := bodyContentStart

	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:insertPos])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[insertPos:])

	return result, nil
}

// insertAtBodyEnd inserts paragraph at the end of document body (before </w:body>)
func insertAtBodyEnd(docXML, paraXML []byte) ([]byte, error) {
	bodyEnd := bytes.Index(docXML, []byte("</w:body>"))
	if bodyEnd == -1 {
		return nil, fmt.Errorf("could not find </w:body> tag")
	}

	insertPos := bodyEnd

	// Find the last <w:sectPr> in the document
	if sectPrPos := bytes.LastIndex(docXML[:bodyEnd], []byte("<w:sectPr")); sectPrPos != -1 {
		// Check if this sectPr is inside a paragraph's properties (<w:pPr>)
		// This happens with section breaks that create new sections
		precedingContent := docXML[:sectPrPos]

		// Find the last <w:pPr> before the sectPr
		lastPPrStart := bytes.LastIndex(precedingContent, []byte("<w:pPr>"))

		if lastPPrStart != -1 {
			// Check if there's a </w:pPr> between lastPPrStart and sectPrPos
			checkRegion := docXML[lastPPrStart:sectPrPos]
			if !bytes.Contains(checkRegion, []byte("</w:pPr>")) {
				// sectPr is inside <w:pPr> (section break in paragraph)
				// Check if there are other paragraphs AFTER this section break paragraph
				// If yes, insert at body end; if no, insert after the section break paragraph

				// Find the end of this section break paragraph
				paraEndPos := bytes.Index(docXML[sectPrPos:bodyEnd], []byte("</w:p>"))
				if paraEndPos != -1 {
					paraEndAbsolute := sectPrPos + paraEndPos + len("</w:p>")

					// Check if there are any other paragraphs between this paragraph end and </w:body>
					afterPara := docXML[paraEndAbsolute:bodyEnd]
					if bytes.Contains(afterPara, []byte("<w:p")) || bytes.Contains(afterPara, []byte("<w:tbl")) {
						// There's content after the section break paragraph, insert at body end
						insertPos = bodyEnd
					} else {
						// No content after section break, insert right after it
						insertPos = paraEndAbsolute
					}
				} else {
					// Fallback: use bodyEnd
					insertPos = bodyEnd
				}
			} else {
				// Normal sectPr (document-level, not in paragraph)
				insertPos = sectPrPos
			}
		} else {
			// No <w:pPr> found, use sectPr position
			insertPos = sectPrPos
		}
	}

	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:insertPos])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[insertPos:])

	return result, nil
}

func findBodyContentStart(docXML []byte) (int, error) {
	bodyStart := bytes.Index(docXML, []byte("<w:body"))
	if bodyStart == -1 {
		return 0, fmt.Errorf("could not find <w:body> tag")
	}

	openTagEnd := bytes.IndexByte(docXML[bodyStart:], '>')
	if openTagEnd == -1 {
		return 0, fmt.Errorf("malformed <w:body> tag")
	}

	return bodyStart + openTagEnd + 1, nil
}

// insertAfterText inserts paragraph after the paragraph containing the anchor text
func insertAfterText(docXML, paraXML []byte, anchorText string) ([]byte, error) {
	_, paraEnd, err := findParagraphRangeByAnchor(docXML, anchorText)
	if err != nil {
		return nil, err
	}

	insertPos := paraEnd

	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:insertPos])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[insertPos:])

	return result, nil
}

// insertBeforeText inserts paragraph before the paragraph containing the anchor text
func insertBeforeText(docXML, paraXML []byte, anchorText string) ([]byte, error) {
	paraStart, _, err := findParagraphRangeByAnchor(docXML, anchorText)
	if err != nil {
		return nil, err
	}

	// Insert before this paragraph
	result := make([]byte, len(docXML)+len(paraXML))
	n := copy(result, docXML[:paraStart])
	n += copy(result[n:], paraXML)
	copy(result[n:], docXML[paraStart:])

	return result, nil
}

func findParagraphRangeByAnchor(docXML []byte, anchorText string) (int, int, error) {
	if anchorText == "" {
		return 0, 0, fmt.Errorf("anchor text cannot be empty")
	}

	normalizedAnchor := normalizeWhitespace(anchorText)

	searchPos := 0
	for {
		paraStart := findNextParagraphStart(docXML, searchPos)
		if paraStart == -1 {
			break
		}

		paraEndRel := bytes.Index(docXML[paraStart:], []byte("</w:p>"))
		if paraEndRel == -1 {
			return 0, 0, fmt.Errorf("could not find paragraph end for anchor search")
		}
		paraEnd := paraStart + paraEndRel + len("</w:p>")

		paragraphXML := docXML[paraStart:paraEnd]
		paragraphText := extractParagraphPlainText(paragraphXML)
		if strings.Contains(paragraphText, anchorText) {
			return paraStart, paraEnd, nil
		}

		if normalizedAnchor != "" && strings.Contains(normalizeWhitespace(paragraphText), normalizedAnchor) {
			return paraStart, paraEnd, nil
		}

		searchPos = paraEnd
	}

	return 0, 0, fmt.Errorf("anchor text %q not found in document", anchorText)
}

func findNextParagraphStart(docXML []byte, start int) int {
	for {
		idx := bytes.Index(docXML[start:], []byte("<w:p"))
		if idx == -1 {
			return -1
		}
		idx += start

		next := idx + len("<w:p")
		if next < len(docXML) {
			ch := docXML[next]
			if ch == '>' || ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r' {
				return idx
			}
		}

		start = idx + len("<w:p")
	}
}

func extractParagraphPlainText(paragraphXML []byte) string {
	var out strings.Builder
	searchPos := 0

	for {
		tStart := findNextWordTagStart(paragraphXML, searchPos, "t")
		tabStart := findNextWordTagStart(paragraphXML, searchPos, "tab")
		brStart := findNextWordTagStart(paragraphXML, searchPos, "br")

		next, kind := nextXMLTokenAbsolute(tStart, tabStart, brStart)
		if next == -1 {
			break
		}

		tStart = next
		if kind == "tab" || kind == "br" {
			out.WriteByte(' ')
			tokenEndRel := bytes.IndexByte(paragraphXML[tStart:], '>')
			if tokenEndRel == -1 {
				break
			}
			searchPos = tStart + tokenEndRel + 1
			continue
		}

		tOpenEndRel := bytes.IndexByte(paragraphXML[tStart:], '>')
		if tOpenEndRel == -1 {
			break
		}
		textStart := tStart + tOpenEndRel + 1

		tCloseRel := bytes.Index(paragraphXML[textStart:], []byte("</w:t>"))
		if tCloseRel == -1 {
			break
		}
		textEnd := textStart + tCloseRel

		out.WriteString(xmlUnescape(string(paragraphXML[textStart:textEnd])))
		searchPos = textEnd + len("</w:t>")
	}

	return out.String()
}

func nextXMLTokenAbsolute(tPos, tabPos, brPos int) (int, string) {
	next := -1
	kind := ""
	set := func(pos int, tokenKind string) {
		if pos == -1 {
			return
		}
		if next == -1 || pos < next {
			next = pos
			kind = tokenKind
		}
	}

	set(tPos, "text")
	set(tabPos, "tab")
	set(brPos, "br")

	return next, kind
}

func findNextWordTagStart(docXML []byte, start int, tag string) int {
	needle := []byte("<w:" + tag)
	for {
		idx := bytes.Index(docXML[start:], needle)
		if idx == -1 {
			return -1
		}
		idx += start

		next := idx + len(needle)
		if next < len(docXML) {
			ch := docXML[next]
			if ch == '>' || ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r' || ch == '/' {
				return idx
			}
		}

		start = idx + len(needle)
	}
}

func normalizeWhitespace(s string) string {
	return strings.Join(strings.Fields(s), " ")
}

// xmlUnescape decodes XML/HTML character references, including named entities
// (&amp; &lt; &gt; &quot; &apos;) and numeric references (&#160; &#x00A0;).
// It delegates to html.UnescapeString which handles the full HTML5 entity table;
// all named entities defined in XML 1.0 are a strict subset of that table.
func xmlUnescape(s string) string {
	return html.UnescapeString(s)
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
		KeepNext: true, // Prevent headings from being orphaned at the bottom of a page
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

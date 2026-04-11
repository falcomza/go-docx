package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// FootnoteOptions defines options for inserting a footnote
type FootnoteOptions struct {
	// Text is the footnote content
	Text string

	// Anchor is the text in the document after which the footnote reference is placed
	Anchor string
}

// EndnoteOptions defines options for inserting an endnote
type EndnoteOptions struct {
	// Text is the endnote content
	Text string

	// Anchor is the text in the document after which the endnote reference is placed
	Anchor string
}

// InsertFootnote adds a footnote to the document.
// The footnote reference marker is placed at the end of the paragraph
// containing the anchor text.
func (u *Updater) InsertFootnote(opts FootnoteOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Text == "" {
		return fmt.Errorf("footnote text cannot be empty")
	}
	if opts.Anchor == "" {
		return fmt.Errorf("anchor text cannot be empty")
	}

	// Ensure footnotes.xml exists and get next footnote ID
	footnoteID, err := u.ensureFootnotesXML()
	if err != nil {
		return fmt.Errorf("ensure footnotes.xml: %w", err)
	}

	if err := u.ensureNoteReferenceStyles(); err != nil {
		return fmt.Errorf("ensure note reference styles: %w", err)
	}

	// Add the footnote content to footnotes.xml
	if err := u.addFootnoteContent(footnoteID, opts.Text); err != nil {
		return fmt.Errorf("add footnote content: %w", err)
	}

	// Insert footnote reference in document.xml
	if err := u.insertNoteReference(opts.Anchor, footnoteID, "footnote"); err != nil {
		return fmt.Errorf("insert footnote reference: %w", err)
	}

	return nil
}

// InsertEndnote adds an endnote to the document.
// The endnote reference marker is placed at the end of the paragraph
// containing the anchor text.
func (u *Updater) InsertEndnote(opts EndnoteOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Text == "" {
		return fmt.Errorf("endnote text cannot be empty")
	}
	if opts.Anchor == "" {
		return fmt.Errorf("anchor text cannot be empty")
	}

	// Ensure endnotes.xml exists and get next endnote ID
	endnoteID, err := u.ensureEndnotesXML()
	if err != nil {
		return fmt.Errorf("ensure endnotes.xml: %w", err)
	}

	if err := u.ensureNoteReferenceStyles(); err != nil {
		return fmt.Errorf("ensure note reference styles: %w", err)
	}

	// Add the endnote content to endnotes.xml
	if err := u.addEndnoteContent(endnoteID, opts.Text); err != nil {
		return fmt.Errorf("add endnote content: %w", err)
	}

	// Insert endnote reference in document.xml
	if err := u.insertNoteReference(opts.Anchor, endnoteID, "endnote"); err != nil {
		return fmt.Errorf("insert endnote reference: %w", err)
	}

	return nil
}

// ensureFootnotesXML creates footnotes.xml if it doesn't exist and returns the next available ID.
func (u *Updater) ensureFootnotesXML() (int, error) {
	fnPath := filepath.Join(u.tempDir, "word", "footnotes.xml")

	if _, err := os.Stat(fnPath); os.IsNotExist(err) {
		// Create initial footnotes.xml with separator footnotes
		content := generateInitialFootnotesXML()
		if err := os.WriteFile(fnPath, content, 0o644); err != nil {
			return 0, fmt.Errorf("write footnotes.xml: %w", err)
		}

		// Add relationship
		if err := u.addNoteRelationship("footnotes.xml", "footnotes"); err != nil {
			return 0, fmt.Errorf("add footnotes relationship: %w", err)
		}

		// Add content type
		if err := u.addNoteContentType("footnotes.xml", "footnotes"); err != nil {
			return 0, fmt.Errorf("add footnotes content type: %w", err)
		}

		return 1, nil
	}

	// Read existing file and find the next available ID
	raw, err := os.ReadFile(fnPath)
	if err != nil {
		return 0, fmt.Errorf("read footnotes.xml: %w", err)
	}

	return getNextNoteID(raw, "footnote"), nil
}

// ensureEndnotesXML creates endnotes.xml if it doesn't exist and returns the next available ID.
func (u *Updater) ensureEndnotesXML() (int, error) {
	enPath := filepath.Join(u.tempDir, "word", "endnotes.xml")

	if _, err := os.Stat(enPath); os.IsNotExist(err) {
		content := generateInitialEndnotesXML()
		if err := os.WriteFile(enPath, content, 0o644); err != nil {
			return 0, fmt.Errorf("write endnotes.xml: %w", err)
		}

		if err := u.addNoteRelationship("endnotes.xml", "endnotes"); err != nil {
			return 0, fmt.Errorf("add endnotes relationship: %w", err)
		}

		if err := u.addNoteContentType("endnotes.xml", "endnotes"); err != nil {
			return 0, fmt.Errorf("add endnotes content type: %w", err)
		}

		return 1, nil
	}

	raw, err := os.ReadFile(enPath)
	if err != nil {
		return 0, fmt.Errorf("read endnotes.xml: %w", err)
	}

	return getNextNoteID(raw, "endnote"), nil
}

// generateInitialFootnotesXML creates a new footnotes.xml with required separator footnotes
func generateInitialFootnotesXML() []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `)
	buf.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString("\n")

	// Separator footnote (required, id=-1)
	buf.WriteString(`<w:footnote w:type="separator" w:id="-1">`)
	buf.WriteString(`<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`)
	buf.WriteString(`<w:r><w:separator/></w:r></w:p>`)
	buf.WriteString(`</w:footnote>`)
	buf.WriteString("\n")

	// Continuation separator footnote (required, id=0)
	buf.WriteString(`<w:footnote w:type="continuationSeparator" w:id="0">`)
	buf.WriteString(`<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`)
	buf.WriteString(`<w:r><w:continuationSeparator/></w:r></w:p>`)
	buf.WriteString(`</w:footnote>`)
	buf.WriteString("\n")

	buf.WriteString(`</w:footnotes>`)

	return buf.Bytes()
}

// generateInitialEndnotesXML creates a new endnotes.xml with required separator endnotes
func generateInitialEndnotesXML() []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `)
	buf.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString("\n")

	buf.WriteString(`<w:endnote w:type="separator" w:id="-1">`)
	buf.WriteString(`<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`)
	buf.WriteString(`<w:r><w:separator/></w:r></w:p>`)
	buf.WriteString(`</w:endnote>`)
	buf.WriteString("\n")

	buf.WriteString(`<w:endnote w:type="continuationSeparator" w:id="0">`)
	buf.WriteString(`<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`)
	buf.WriteString(`<w:r><w:continuationSeparator/></w:r></w:p>`)
	buf.WriteString(`</w:endnote>`)
	buf.WriteString("\n")

	buf.WriteString(`</w:endnotes>`)

	return buf.Bytes()
}

// addFootnoteContent adds a footnote entry to footnotes.xml
func (u *Updater) addFootnoteContent(id int, text string) error {
	fnPath := filepath.Join(u.tempDir, "word", "footnotes.xml")
	raw, err := os.ReadFile(fnPath)
	if err != nil {
		return fmt.Errorf("read footnotes.xml: %w", err)
	}

	footnoteXML := generateFootnoteEntry(id, text)

	// Insert before </w:footnotes>
	closeTag := []byte("</w:footnotes>")
	closeIdx := bytes.LastIndex(raw, closeTag)
	if closeIdx == -1 {
		return fmt.Errorf("could not find </w:footnotes> tag")
	}

	result := make([]byte, 0, len(raw)+len(footnoteXML)+1)
	result = append(result, raw[:closeIdx]...)
	result = append(result, footnoteXML...)
	result = append(result, '\n')
	result = append(result, raw[closeIdx:]...)

	if err := os.WriteFile(fnPath, result, 0o644); err != nil {
		return fmt.Errorf("write footnotes.xml: %w", err)
	}

	return nil
}

// addEndnoteContent adds an endnote entry to endnotes.xml
func (u *Updater) addEndnoteContent(id int, text string) error {
	enPath := filepath.Join(u.tempDir, "word", "endnotes.xml")
	raw, err := os.ReadFile(enPath)
	if err != nil {
		return fmt.Errorf("read endnotes.xml: %w", err)
	}

	endnoteXML := generateEndnoteEntry(id, text)

	closeTag := []byte("</w:endnotes>")
	closeIdx := bytes.LastIndex(raw, closeTag)
	if closeIdx == -1 {
		return fmt.Errorf("could not find </w:endnotes> tag")
	}

	result := make([]byte, 0, len(raw)+len(endnoteXML)+1)
	result = append(result, raw[:closeIdx]...)
	result = append(result, endnoteXML...)
	result = append(result, '\n')
	result = append(result, raw[closeIdx:]...)

	if err := os.WriteFile(enPath, result, 0o644); err != nil {
		return fmt.Errorf("write endnotes.xml: %w", err)
	}

	return nil
}

// generateFootnoteEntry creates the XML for a single footnote
func generateFootnoteEntry(id int, text string) []byte {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`<w:footnote w:id="%d">`, id))
	buf.WriteString("<w:p>")
	buf.WriteString(`<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>`)

	// Footnote reference marker (the superscript number in the footnote area)
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>`)
	buf.WriteString("<w:footnoteRef/>")
	buf.WriteString("</w:r>")

	// Space + text
	buf.WriteString("<w:r>")
	buf.WriteString(fmt.Sprintf(`<w:t xml:space="preserve"> %s</w:t>`, xmlEscape(text)))
	buf.WriteString("</w:r>")

	buf.WriteString("</w:p>")
	buf.WriteString("</w:footnote>")

	return buf.Bytes()
}

// generateEndnoteEntry creates the XML for a single endnote
func generateEndnoteEntry(id int, text string) []byte {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`<w:endnote w:id="%d">`, id))
	buf.WriteString("<w:p>")
	buf.WriteString(`<w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>`)

	buf.WriteString("<w:r>")
	buf.WriteString(`<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>`)
	buf.WriteString("<w:endnoteRef/>")
	buf.WriteString("</w:r>")

	buf.WriteString("<w:r>")
	buf.WriteString(fmt.Sprintf(`<w:t xml:space="preserve"> %s</w:t>`, xmlEscape(text)))
	buf.WriteString("</w:r>")

	buf.WriteString("</w:p>")
	buf.WriteString("</w:endnote>")

	return buf.Bytes()
}

// insertNoteReference inserts a footnote or endnote reference into document.xml
// at the end of the paragraph containing the anchor text.
func (u *Updater) insertNoteReference(anchor string, noteID int, noteType string) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Find the paragraph containing the anchor text
	_, paraEnd, err := findParagraphRangeByAnchor(raw, anchor)
	if err != nil {
		return fmt.Errorf("find anchor: %w", err)
	}

	// Build the reference run XML
	var refXML string
	if noteType == "footnote" {
		refXML = fmt.Sprintf(
			`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>`+
				`<w:footnoteReference w:id="%d"/></w:r>`, noteID)
	} else {
		refXML = fmt.Sprintf(
			`<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>`+
				`<w:endnoteReference w:id="%d"/></w:r>`, noteID)
	}

	// Insert the reference run before </w:p>
	// paraEnd points to the end of "</w:p>", so we insert before that tag
	insertPos := paraEnd - len("</w:p>")

	result := make([]byte, 0, len(raw)+len(refXML))
	result = append(result, raw[:insertPos]...)
	result = append(result, []byte(refXML)...)
	result = append(result, raw[insertPos:]...)

	if err := os.WriteFile(docPath, result, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// getNextNoteID finds the next available note ID in a footnotes/endnotes XML file
func getNextNoteID(raw []byte, noteType string) int {
	pattern := regexp.MustCompile(fmt.Sprintf(`<w:%s[^>]*w:id="(\d+)"`, noteType))
	matches := pattern.FindAllSubmatch(raw, -1)

	maxID := 0
	for _, match := range matches {
		if len(match) > 1 {
			id, err := strconv.Atoi(string(match[1]))
			if err != nil {
				continue
			}
			if id > maxID {
				maxID = id
			}
		}
	}

	return maxID + 1
}

// addNoteRelationship adds a relationship for footnotes or endnotes
func (u *Updater) addNoteRelationship(filename, relType string) error {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read rels: %w", err)
	}

	content := string(raw)

	// Check if already exists
	if strings.Contains(content, filename) {
		return nil
	}

	relID, err := getNextRelIDFromFile(relsPath)
	if err != nil {
		return fmt.Errorf("get next rel ID: %w", err)
	}

	newRel := fmt.Sprintf(
		`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s" Target="%s"/>`,
		relID, relType, filename,
	)

	content = strings.Replace(content, "</Relationships>", newRel+"</Relationships>", 1)

	if err := os.WriteFile(relsPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write rels: %w", err)
	}

	return nil
}

// ensureNoteReferenceStyles guarantees that the FootnoteReference and EndnoteReference
// character styles exist in styles.xml with <w:vertAlign w:val="superscript"/>.
//
// Word only renders in-text note reference markers as superscript when the character
// style explicitly carries that property. LibreOffice hard-codes the behaviour, but
// Word falls back to normal-size text when the style is absent or lacks vertAlign.
func (u *Updater) ensureNoteReferenceStyles() error {
	if u.noteRefStylesEnsured {
		return nil
	}

	stylesPath := filepath.Join(u.tempDir, "word", "styles.xml")
	raw, err := os.ReadFile(stylesPath)
	wasAbsent := os.IsNotExist(err)
	if err != nil && !wasAbsent {
		return fmt.Errorf("read styles.xml: %w", err)
	}

	// Collect XML for any missing note reference styles in a single pass.
	var missing []byte
	for _, ns := range [2][2]string{
		{"FootnoteReference", "footnote reference"},
		{"EndnoteReference", "endnote reference"},
	} {
		id, name := ns[0], ns[1]
		if strings.Contains(string(raw), `w:styleId="`+id+`"`) {
			continue
		}
		missing = append(missing, fmt.Sprintf(
			`<w:style w:type="character" w:styleId="%s">`+
				`<w:name w:val="%s"/>`+
				`<w:basedOn w:val="DefaultParagraphFont"/>`+
				`<w:uiPriority w:val="99"/>`+
				`<w:semiHidden/>`+
				`<w:unhideWhenUsed/>`+
				`<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`+
				`</w:style>`,
			id, name,
		)...)
	}

	if len(missing) == 0 {
		u.noteRefStylesEnsured = true
		return nil
	}

	var updated []byte
	if wasAbsent {
		updated = generateStylesDocument(missing)
	} else {
		updated, err = injectStyle(raw, missing)
		if err != nil {
			return fmt.Errorf("inject note reference styles: %w", err)
		}
	}

	if err := atomicWriteFile(stylesPath, updated, 0o644); err != nil {
		return fmt.Errorf("write styles.xml: %w", err)
	}
	if wasAbsent {
		if err := u.ensureStylesRelationship(); err != nil {
			return fmt.Errorf("ensure styles relationship: %w", err)
		}
	}

	u.noteRefStylesEnsured = true
	return nil
}

// addNoteContentType adds a content type for footnotes or endnotes
func (u *Updater) addNoteContentType(filename, noteType string) error {
	ctPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	raw, err := os.ReadFile(ctPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	content := string(raw)

	if strings.Contains(content, filename) {
		return nil
	}

	contentType := fmt.Sprintf(
		"application/vnd.openxmlformats-officedocument.wordprocessingml.%s+xml", noteType)

	override := fmt.Sprintf(
		`<Override PartName="/word/%s" ContentType="%s"/>`,
		filename, contentType,
	)

	content = strings.Replace(content, "</Types>", override+"</Types>", 1)

	if err := os.WriteFile(ctPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write content types: %w", err)
	}

	return nil
}

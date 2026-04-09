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

// List numbering IDs
const (
	BulletListNumID   = 1 // Numbering ID for bullet lists
	NumberedListNumID = 2 // Numbering ID for numbered lists
)

var (
	docxUpdateBulletNumIDPattern   = regexp.MustCompile(`DOCXUPDATE_BULLET_NUMID:(\d+)`)
	docxUpdateNumberedNumIDPattern = regexp.MustCompile(`DOCXUPDATE_NUMBERED_NUMID:(\d+)`)
	docxUpdateHeadingNumIDPattern  = regexp.MustCompile(`DOCXUPDATE_HEADING_NUMID:(\d+)`)
	numIDPattern                   = regexp.MustCompile(`w:numId="(\d+)"`)
	abstractNumIDPattern           = regexp.MustCompile(`w:abstractNumId="(\d+)"`)
	// abstractNumRefPattern matches the child element <w:abstractNumId w:val="N"/> inside a <w:num> block.
	abstractNumRefPattern = regexp.MustCompile(`<w:abstractNumId[^>]+w:val="(\d+)"`)
)

// ensureNumberingXML ensures numbering.xml exists with bullet and numbered list support.
// It returns the final content of numbering.xml so callers can avoid a second read.
func (u *Updater) ensureNumberingXML() ([]byte, error) {
	numberingPath := filepath.Join(u.tempDir, "word", "numbering.xml")

	var finalContent []byte
	if data, err := os.ReadFile(numberingPath); err == nil {
		content := string(data)

		if bulletID, numberedID, ok := extractDocxUpdateNumberingIDs(content); ok {
			u.setListNumberingIDs(bulletID, numberedID)
			finalContent = data
		} else {
			updated, bulletID, numberedID, appendErr := appendDocxUpdateNumberingDefinitions(content)
			if appendErr != nil {
				return nil, fmt.Errorf("append numbering definitions: %w", appendErr)
			}
			if err := atomicWriteFile(numberingPath, []byte(updated), 0o644); err != nil {
				return nil, fmt.Errorf("write numbering.xml: %w", err)
			}
			u.setListNumberingIDs(bulletID, numberedID)
			finalContent = []byte(updated)
		}
	} else if !os.IsNotExist(err) {
		return nil, fmt.Errorf("read numbering.xml: %w", err)
	} else {
		numberingXML := generateNumberingXML()
		if err := atomicWriteFile(numberingPath, []byte(numberingXML), 0o644); err != nil {
			return nil, fmt.Errorf("write numbering.xml: %w", err)
		}
		u.setListNumberingIDs(BulletListNumID, NumberedListNumID)
		finalContent = []byte(numberingXML)
	}

	// Update content types if needed
	if err := u.ensureNumberingContentType(); err != nil {
		return nil, fmt.Errorf("update content types: %w", err)
	}

	// Update document.xml.rels if needed
	if err := u.ensureNumberingRelationship(); err != nil {
		return nil, fmt.Errorf("update relationships: %w", err)
	}

	// Ensure ListParagraph style is defined in styles.xml
	if err := u.ensureListParagraphStyle(); err != nil {
		return nil, fmt.Errorf("ensure ListParagraph style: %w", err)
	}

	return finalContent, nil
}

// generateNumberingXML creates a complete numbering.xml with bullet and numbered list definitions
func generateNumberingXML() string {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	buf.WriteString(`<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` + "\n")
	buf.WriteString(`             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` + "\n")
	buf.WriteString(`             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` + "\n")
	buf.WriteString(`             mc:Ignorable="w14">` + "\n")
	buf.WriteString(generateDocxUpdateNumberingDefinitions(0, 1, BulletListNumID, NumberedListNumID))
	buf.WriteString("\n</w:numbering>")

	return buf.String()
}

func generateDocxUpdateNumberingDefinitions(bulletAbstractID, numberedAbstractID, bulletNumID, numberedNumID int) string {
	var buf bytes.Buffer

	bulletSymbols := []string{"●", "○", "■", "●", "○", "■", "●", "○", "■"}
	bulletFonts := []string{"Symbol", "Courier New", "Wingdings", "", "", "", "", "", ""}

	buf.WriteString("\n  <!-- Abstract Numbering Definition for Bullets -->\n")
	fmt.Fprintf(&buf, "  <w:abstractNum w:abstractNumId=\"%d\">\n", bulletAbstractID)
	buf.WriteString("    <w:multiLevelType w:val=\"hybridMultilevel\"/>\n")
	for level := 0; level <= 8; level++ {
		left := 720 * (level + 1)
		fmt.Fprintf(&buf, "    <w:lvl w:ilvl=\"%d\">\n", level)
		buf.WriteString("      <w:start w:val=\"1\"/>\n")
		buf.WriteString("      <w:numFmt w:val=\"bullet\"/>\n")
		fmt.Fprintf(&buf, "      <w:lvlText w:val=\"%s\"/>\n", bulletSymbols[level])
		buf.WriteString("      <w:lvlJc w:val=\"left\"/>\n")
		buf.WriteString("      <w:pPr>\n")
		fmt.Fprintf(&buf, "        <w:ind w:left=\"%d\" w:hanging=\"360\"/>\n", left)
		buf.WriteString("      </w:pPr>\n")
		if font := bulletFonts[level]; font != "" {
			buf.WriteString("      <w:rPr>\n")
			fmt.Fprintf(&buf, "        <w:rFonts w:ascii=\"%s\" w:hAnsi=\"%s\" w:hint=\"default\"/>\n", font, font)
			buf.WriteString("      </w:rPr>\n")
		}
		buf.WriteString("    </w:lvl>\n")
	}
	buf.WriteString("  </w:abstractNum>\n")

	numberedFormats := []string{
		"decimal", "decimal", "decimal", "decimal", "decimal",
		"decimal", "decimal", "decimal", "decimal",
	}

	numberedTexts := []string{
		"%1.",
		"%1.%2.",
		"%1.%2.%3.",
		"%1.%2.%3.%4.",
		"%1.%2.%3.%4.%5.",
		"%1.%2.%3.%4.%5.%6.",
		"%1.%2.%3.%4.%5.%6.%7.",
		"%1.%2.%3.%4.%5.%6.%7.%8.",
		"%1.%2.%3.%4.%5.%6.%7.%8.%9.",
	}

	buf.WriteString("\n  <!-- Abstract Numbering Definition for Numbered Lists -->\n")
	fmt.Fprintf(&buf, "  <w:abstractNum w:abstractNumId=\"%d\">\n", numberedAbstractID)
	buf.WriteString("    <w:multiLevelType w:val=\"hybridMultilevel\"/>\n")
	for level := 0; level <= 8; level++ {
		left := 720 * (level + 1)
		fmt.Fprintf(&buf, "    <w:lvl w:ilvl=\"%d\">\n", level)
		buf.WriteString("      <w:start w:val=\"1\"/>\n")
		fmt.Fprintf(&buf, "      <w:numFmt w:val=\"%s\"/>\n", numberedFormats[level])
		fmt.Fprintf(&buf, "      <w:lvlText w:val=\"%s\"/>\n", numberedTexts[level])
		buf.WriteString("      <w:lvlJc w:val=\"left\"/>\n")
		buf.WriteString("      <w:pPr>\n")
		fmt.Fprintf(&buf, "        <w:ind w:left=\"%d\" w:hanging=\"360\"/>\n", left)
		buf.WriteString("      </w:pPr>\n")
		buf.WriteString("    </w:lvl>\n")
	}
	buf.WriteString("  </w:abstractNum>\n")

	buf.WriteString("\n")
	fmt.Fprintf(&buf, "  <!-- DOCXUPDATE_BULLET_NUMID:%d -->\n", bulletNumID)
	fmt.Fprintf(&buf, "  <w:num w:numId=\"%d\">\n", bulletNumID)
	fmt.Fprintf(&buf, "    <w:abstractNumId w:val=\"%d\"/>\n", bulletAbstractID)
	buf.WriteString("  </w:num>\n")

	fmt.Fprintf(&buf, "  <!-- DOCXUPDATE_NUMBERED_NUMID:%d -->\n", numberedNumID)
	fmt.Fprintf(&buf, "  <w:num w:numId=\"%d\">\n", numberedNumID)
	fmt.Fprintf(&buf, "    <w:abstractNumId w:val=\"%d\"/>\n", numberedAbstractID)
	buf.WriteString("  </w:num>\n")

	return buf.String()
}

func extractDocxUpdateNumberingIDs(content string) (int, int, bool) {
	bulletMatch := docxUpdateBulletNumIDPattern.FindStringSubmatch(content)
	numberedMatch := docxUpdateNumberedNumIDPattern.FindStringSubmatch(content)
	if len(bulletMatch) < 2 || len(numberedMatch) < 2 {
		return 0, 0, false
	}

	bulletID, err := strconv.Atoi(bulletMatch[1])
	if err != nil {
		return 0, 0, false
	}
	numberedID, err := strconv.Atoi(numberedMatch[1])
	if err != nil {
		return 0, 0, false
	}
	if bulletID <= 0 || numberedID <= 0 {
		return 0, 0, false
	}

	return bulletID, numberedID, true
}

// appendToNumberingXML inserts generated XML before the closing </w:numbering> tag.
// generateFn receives the next available abstractNumId and numId and returns the XML to insert.
// It returns the updated content and the two IDs that were allocated (abstractID, numID).
func appendToNumberingXML(content string, generateFn func(abstractID, numID int) string) (string, int, int, error) {
	closingTag := "</w:numbering>"
	insertPos := strings.LastIndex(content, closingTag)
	if insertPos == -1 {
		return "", 0, 0, fmt.Errorf("invalid numbering.xml: missing </w:numbering>")
	}

	abstractID := findMaxXMLAttributeInt(content, abstractNumIDPattern) + 1
	numID := findMaxXMLAttributeInt(content, numIDPattern) + 1

	definitions := generateFn(abstractID, numID)
	updated := content[:insertPos] + definitions + "\n" + content[insertPos:]
	return updated, abstractID, numID, nil
}

func appendDocxUpdateNumberingDefinitions(content string) (string, int, int, error) {
	maxAbstractNumID := findMaxXMLAttributeInt(content, abstractNumIDPattern)
	maxNumID := findMaxXMLAttributeInt(content, numIDPattern)

	bulletAbstractID := maxAbstractNumID + 1
	numberedAbstractID := maxAbstractNumID + 2
	bulletNumID := maxNumID + 1
	numberedNumID := maxNumID + 2

	gen := func(_, _ int) string {
		return generateDocxUpdateNumberingDefinitions(bulletAbstractID, numberedAbstractID, bulletNumID, numberedNumID)
	}
	updated, _, _, err := appendToNumberingXML(content, gen)
	if err != nil {
		return "", 0, 0, err
	}
	return updated, bulletNumID, numberedNumID, nil
}

func findMaxXMLAttributeInt(content string, pattern *regexp.Regexp) int {
	maxValue := 0
	matches := pattern.FindAllStringSubmatch(content, -1)
	for _, match := range matches {
		if len(match) < 2 {
			continue
		}
		value, err := strconv.Atoi(match[1])
		if err != nil {
			continue
		}
		if value > maxValue {
			maxValue = value
		}
	}
	return maxValue
}

func (u *Updater) setListNumberingIDs(bulletID, numberedID int) {
	u.bulletListNumID = bulletID
	u.numberedListNumID = numberedID
}

// allocateRestartNumIDInContent appends a new <w:num> entry to an in-memory
// numbering.xml content string, returning the new numId and updated content.
// level is clamped to [0, 8]. numberedNumID is the numId of the base numbered list.
func allocateRestartNumIDInContent(content string, numberedNumID, level int) (int, string, error) {
	abstractID := findAbstractNumIDForNum(content, numberedNumID)
	if abstractID < 0 {
		return 0, "", fmt.Errorf("could not find abstractNumId for numbered list numId=%d", numberedNumID)
	}

	maxNumID := findMaxXMLAttributeInt(content, numIDPattern)
	newNumID := maxNumID + 1

	level = min(max(level, 0), 8)

	newNum := fmt.Sprintf(
		"\n  <w:num w:numId=\"%d\">\n    <w:abstractNumId w:val=\"%d\"/>\n    <w:lvlOverride w:ilvl=\"%d\">\n      <w:startOverride w:val=\"1\"/>\n    </w:lvlOverride>\n  </w:num>",
		newNumID, abstractID, level,
	)

	closingTag := "</w:numbering>"
	insertPos := strings.LastIndex(content, closingTag)
	if insertPos == -1 {
		return 0, "", fmt.Errorf("invalid numbering.xml: missing </w:numbering>")
	}

	updated := content[:insertPos] + newNum + "\n" + content[insertPos:]
	return newNumID, updated, nil
}

// AllocateRestartNumID appends a new <w:num> entry to numbering.xml that references
// the same abstractNumId as the current numbered list but adds a
// <w:lvlOverride><w:startOverride w:val="1"/></w:lvlOverride> for the given level.
// It returns the newly allocated numId. ensureNumberingXML must have been called first.
//
// Use this to get a fresh numId for a new numbered-list sequence. Pass the returned
// numId as ParagraphOptions.OverrideNumID for every item in the sequence so that all
// items share the same counter (which restarts from 1) instead of accumulating on the
// shared base numberedListNumID.
func (u *Updater) AllocateRestartNumID(level int) (int, error) {
	return u.allocateRestartNumID(level)
}

// allocateRestartNumID is the internal implementation; see AllocateRestartNumID for docs.
func (u *Updater) allocateRestartNumID(level int) (int, error) {
	numberingPath := filepath.Join(u.tempDir, "word", "numbering.xml")
	data, err := os.ReadFile(numberingPath)
	if err != nil {
		return 0, fmt.Errorf("read numbering.xml: %w", err)
	}

	ids := u.getListNumberingIDs()
	newNumID, updated, err := allocateRestartNumIDInContent(string(data), ids.numberedNumID, level)
	if err != nil {
		return 0, err
	}

	if err := atomicWriteFile(numberingPath, []byte(updated), 0o644); err != nil {
		return 0, fmt.Errorf("write numbering.xml: %w", err)
	}

	return newNumID, nil
}

// findAbstractNumIDForNum returns the abstractNumId referenced by the given numId,
// or -1 if not found.
func findAbstractNumIDForNum(content string, numID int) int {
	// Locate the exact opening tag <w:num w:numId="N"> to avoid partial-number matches.
	numBlock := fmt.Sprintf(`<w:num w:numId="%d">`, numID)
	idx := strings.Index(content, numBlock)
	if idx == -1 {
		return -1
	}
	rest := content[idx:]
	closeIdx := strings.Index(rest, "</w:num>")
	if closeIdx == -1 {
		return -1
	}
	block := rest[:closeIdx]
	match := abstractNumRefPattern.FindStringSubmatch(block)
	if len(match) < 2 {
		return -1
	}
	val, err := strconv.Atoi(match[1])
	if err != nil {
		return -1
	}
	return val
}

func (u *Updater) getListNumberingIDs() listNumberingIDs {
	ids := listNumberingIDs{bulletNumID: u.bulletListNumID, numberedNumID: u.numberedListNumID}
	if ids.bulletNumID <= 0 {
		ids.bulletNumID = BulletListNumID
	}
	if ids.numberedNumID <= 0 {
		ids.numberedNumID = NumberedListNumID
	}
	return ids
}

// ensureNumberingContentType adds numbering.xml to [Content_Types].xml if not present
func (u *Updater) ensureNumberingContentType() error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	data, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read [Content_Types].xml: %w", err)
	}

	content := string(data)

	// Check if numbering override already exists
	if strings.Contains(content, `PartName="/word/numbering.xml"`) {
		return nil // Already present
	}

	// Add numbering override before </Types>
	numberingOverride := `  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
	content = strings.Replace(content, "</Types>", numberingOverride+"\n</Types>", 1)

	return atomicWriteFile(contentTypesPath, []byte(content), 0o644)
}

// ensureNumberingRelationship adds numbering.xml relationship to document.xml.rels if not present
func (u *Updater) ensureNumberingRelationship() error {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	data, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read document.xml.rels: %w", err)
	}

	content := string(data)

	// Check if numbering relationship already exists
	if strings.Contains(content, `Target="numbering.xml"`) {
		return nil // Already present
	}

	// Find the next available relationship ID
	relID, err := getNextRelIDFromFile(relsPath)
	if err != nil {
		return fmt.Errorf("find next relationship id: %w", err)
	}

	// Add numbering relationship before </Relationships>
	numberingRel := fmt.Sprintf(`  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`, relID)
	content = strings.Replace(content, "</Relationships>", numberingRel+"\n</Relationships>", 1)

	return atomicWriteFile(relsPath, []byte(content), 0o644)
}

// ensureListParagraphStyle ensures that the ListParagraph style is defined in styles.xml
// This prevents MS Word from showing corruption warnings about missing list styles
func (u *Updater) ensureListParagraphStyle() error {
	stylesPath := filepath.Join(u.tempDir, "word", "styles.xml")

	// Read current styles.xml or create minimal one
	data, err := os.ReadFile(stylesPath)
	if err != nil {
		if !os.IsNotExist(err) {
			return fmt.Errorf("read styles.xml: %w", err)
		}
		// Create minimal styles.xml with ListParagraph style
		content := generateMinimalStylesXMLWithListParagraph()
		if err := atomicWriteFile(stylesPath, []byte(content), 0o644); err != nil {
			return fmt.Errorf("write styles.xml: %w", err)
		}
		// Ensure styles relationship exists
		if err := u.ensureStylesRelationship(); err != nil {
			return fmt.Errorf("ensure styles relationship: %w", err)
		}
		return nil
	}

	content := string(data)

	// Check if ListParagraph style already exists
	if strings.Contains(content, `w:styleId="ListParagraph"`) {
		return nil // Already present
	}

	// Add ListParagraph style before </w:styles>
	listParagraphStyle := generateListParagraphStyleXML()
	closingTag := "</w:styles>"
	insertPos := strings.LastIndex(content, closingTag)
	if insertPos == -1 {
		return fmt.Errorf("invalid styles.xml: missing </w:styles>")
	}

	updated := content[:insertPos] + listParagraphStyle + "\n" + content[insertPos:]
	return atomicWriteFile(stylesPath, []byte(updated), 0o644)
}

// generateListParagraphStyleXML creates the XML for the ListParagraph style
func generateListParagraphStyleXML() string {
	var buf bytes.Buffer
	buf.WriteString("\n  <!-- List Paragraph Style - Required for proper list formatting -->\n")
	buf.WriteString(`  <w:style w:type="paragraph" w:styleId="ListParagraph">`)
	buf.WriteString("\n")
	buf.WriteString(`    <w:name w:val="List Paragraph"/>`)
	buf.WriteString("\n")
	buf.WriteString(`    <w:basedOn w:val="Normal"/>`)
	buf.WriteString("\n")
	buf.WriteString(`    <w:qFormat/>`)
	buf.WriteString("\n")
	buf.WriteString(`    <w:pPr>`)
	buf.WriteString("\n")
	buf.WriteString(`      <w:ind w:left="720"/>`)
	buf.WriteString("\n")
	buf.WriteString(`      <w:contextualSpacing/>`)
	buf.WriteString("\n")
	buf.WriteString(`    </w:pPr>`)
	buf.WriteString("\n")
	buf.WriteString(`  </w:style>`)
	buf.WriteString("\n")
	return buf.String()
}

// generateMinimalStylesXMLWithListParagraph creates a minimal styles.xml with ListParagraph style
func generateMinimalStylesXMLWithListParagraph() string {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)
	buf.WriteString("\n")
	buf.WriteString(`  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">`)
	buf.WriteString("\n")
	buf.WriteString(`    <w:name w:val="Normal"/>`)
	buf.WriteString("\n")
	buf.WriteString(`  </w:style>`)
	buf.WriteString("\n")
	buf.WriteString(generateListParagraphStyleXML())
	buf.WriteString("</w:styles>")
	return buf.String()
}

// generateHeadingOutlineXML creates the abstractNum + num definitions for heading outline
// numbering (1 / 1.1 / 1.1.1 …). The DOCXUPDATE_HEADING_NUMID comment acts as the
// idempotency marker so the definitions are only appended once per document.
func generateHeadingOutlineXML(abstractID, numID int) string {
	var buf bytes.Buffer

	// Precomputed lvlText patterns with trailing dots: "1." / "1.2." / "1.2.3." / …
	lvlTexts := [9]string{
		"%1.",
		"%1.%2.",
		"%1.%2.%3.",
		"%1.%2.%3.%4.",
		"%1.%2.%3.%4.%5.",
		"%1.%2.%3.%4.%5.%6.",
		"%1.%2.%3.%4.%5.%6.%7.",
		"%1.%2.%3.%4.%5.%6.%7.%8.",
		"%1.%2.%3.%4.%5.%6.%7.%8.%9.",
	}

	buf.WriteString("\n  <!-- Abstract Numbering Definition for Heading Outline -->\n")
	fmt.Fprintf(&buf, "  <w:abstractNum w:abstractNumId=\"%d\">\n", abstractID)
	buf.WriteString("    <w:multiLevelType w:val=\"multilevel\"/>\n")
	for level := 0; level <= 8; level++ {
		// OOXML CT_Lvl child order (ECMA-376 §17.9.6): start, numFmt, lvlText, lvlJc, pPr
		// NOTE: <w:pStyle> is intentionally omitted. Linking an abstractNum level to a
		// heading style via <w:pStyle> creates a style-list connection that causes Word
		// to override the indentation of ALL paragraphs using that style — including
		// those already defined in the uploaded template. Since we apply <w:numPr>
		// explicitly on each heading paragraph, the pStyle link is not required for
		// correct cascade numbering and its absence preserves template-defined indentation.
		//
		// <w:ind w:left="0" w:hanging="0"/> is explicitly set so that Word does not apply
		// its internal default indentation (720 * (ilvl+1) twips) for numbered list levels.
		// LibreOffice ignores the missing indent, but Word applies it unless suppressed here.
		// Priority (ECMA-376 §17.3.1): numbering-level pPr < paragraph style < direct fmt,
		// so templates that define their own Heading indent will still take precedence.
		fmt.Fprintf(&buf, "    <w:lvl w:ilvl=\"%d\">\n", level)
		buf.WriteString("      <w:start w:val=\"1\"/>\n")
		buf.WriteString("      <w:numFmt w:val=\"decimal\"/>\n")
		fmt.Fprintf(&buf, "      <w:lvlText w:val=\"%s\"/>\n", lvlTexts[level])
		buf.WriteString("      <w:lvlJc w:val=\"left\"/>\n")
		buf.WriteString("      <w:pPr>\n")
		buf.WriteString("        <w:ind w:left=\"0\" w:hanging=\"0\"/>\n")
		buf.WriteString("      </w:pPr>\n")
		buf.WriteString("    </w:lvl>\n")
	}
	buf.WriteString("  </w:abstractNum>\n")

	fmt.Fprintf(&buf, "\n  <!-- DOCXUPDATE_HEADING_NUMID:%d -->\n", numID)
	fmt.Fprintf(&buf, "  <w:num w:numId=\"%d\">\n", numID)
	fmt.Fprintf(&buf, "    <w:abstractNumId w:val=\"%d\"/>\n", abstractID)
	buf.WriteString("  </w:num>\n")

	return buf.String()
}

// ensureHeadingOutlineNumbering guarantees that a multi-level heading outline
// numbering definition exists in numbering.xml and returns its numId.
// Idempotent: a second call returns the cached ID without re-reading the file.
func (u *Updater) ensureHeadingOutlineNumbering() (int, error) {
	if u.headingNumID > 0 {
		return u.headingNumID, nil
	}

	// Ensure numbering.xml itself exists first; reuse the returned content to avoid a second read.
	data, err := u.ensureNumberingXML()
	if err != nil {
		return 0, fmt.Errorf("ensure numbering XML: %w", err)
	}
	content := string(data)

	// If the marker already exists, just cache and return the ID.
	if match := docxUpdateHeadingNumIDPattern.FindStringSubmatch(content); len(match) >= 2 {
		id, err := strconv.Atoi(match[1])
		if err != nil || id <= 0 {
			return 0, fmt.Errorf("invalid DOCXUPDATE_HEADING_NUMID value: %s", match[1])
		}
		u.headingNumID = id
		return id, nil
	}

	// Append new heading outline definitions.
	updated, _, numID, err := appendToNumberingXML(content, generateHeadingOutlineXML)
	if err != nil {
		return 0, err
	}

	if err := atomicWriteFile(filepath.Join(u.tempDir, "word", "numbering.xml"), []byte(updated), 0o644); err != nil {
		return 0, fmt.Errorf("write numbering.xml: %w", err)
	}

	u.headingNumID = numID
	return numID, nil
}

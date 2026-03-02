package godocx

import (
	"strings"
	"testing"
	"time"
)

// Golden tests compare generated XML against expected output to catch
// unintentional changes in XML generation functions.

func TestGolden_GenerateParagraphXML_Normal(t *testing.T) {
	opts := ParagraphOptions{
		Text:  "Hello world",
		Style: StyleNormal,
	}
	listIDs := listNumberingIDs{bulletNumID: 1, numberedNumID: 2}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	expected := `<w:p><w:pPr></w:pPr><w:r><w:t>Hello world</w:t></w:r></w:p>`
	if result != expected {
		t.Errorf("paragraph XML mismatch\ngot:  %s\nwant: %s", result, expected)
	}
}

func TestGolden_GenerateParagraphXML_Heading(t *testing.T) {
	opts := ParagraphOptions{
		Text:  "My Heading",
		Style: StyleHeading1,
	}
	listIDs := listNumberingIDs{}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	if !strings.Contains(result, `<w:pStyle w:val="Heading1"/>`) {
		t.Errorf("expected Heading1 style in: %s", result)
	}
	if !strings.Contains(result, "My Heading") {
		t.Errorf("expected text in: %s", result)
	}
}

func TestGolden_GenerateParagraphXML_BoldItalic(t *testing.T) {
	opts := ParagraphOptions{
		Text:      "Formatted",
		Style:     StyleNormal,
		Bold:      true,
		Italic:    true,
		Underline: true,
	}
	listIDs := listNumberingIDs{}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	expected := `<w:p><w:pPr></w:pPr><w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>Formatted</w:t></w:r></w:p>`
	if result != expected {
		t.Errorf("formatted paragraph XML mismatch\ngot:  %s\nwant: %s", result, expected)
	}
}

func TestGolden_GenerateParagraphXML_Alignment(t *testing.T) {
	opts := ParagraphOptions{
		Text:      "Centered",
		Style:     StyleNormal,
		Alignment: ParagraphAlignCenter,
	}
	listIDs := listNumberingIDs{}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	if !strings.Contains(result, `<w:jc w:val="center"/>`) {
		t.Errorf("expected center alignment in: %s", result)
	}
}

func TestGolden_GenerateParagraphXML_BulletList(t *testing.T) {
	opts := ParagraphOptions{
		Text:      "Bullet item",
		Style:     StyleNormal,
		ListType:  ListTypeBullet,
		ListLevel: 0,
	}
	listIDs := listNumberingIDs{bulletNumID: 5, numberedNumID: 6}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	if !strings.Contains(result, `<w:ilvl w:val="0"/>`) {
		t.Errorf("expected ilvl in: %s", result)
	}
	if !strings.Contains(result, `<w:numId w:val="5"/>`) {
		t.Errorf("expected numId 5 in: %s", result)
	}
}

func TestGolden_GenerateParagraphXML_LineBreaks(t *testing.T) {
	opts := ParagraphOptions{
		Text:  "Line1\nLine2\tTabbed",
		Style: StyleNormal,
	}
	listIDs := listNumberingIDs{}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	if !strings.Contains(result, "<w:br/>") {
		t.Errorf("expected <w:br/> for newline in: %s", result)
	}
	if !strings.Contains(result, "<w:tab/>") {
		t.Errorf("expected <w:tab/> for tab in: %s", result)
	}
}

func TestGolden_GenerateParagraphXML_XMLEscaping(t *testing.T) {
	opts := ParagraphOptions{
		Text:  `Text with <special> & "chars"`,
		Style: StyleNormal,
	}
	listIDs := listNumberingIDs{}

	result := string(generateParagraphXML(opts, listIDs, map[string]string{}))

	if !strings.Contains(result, "&lt;special&gt;") {
		t.Errorf("expected escaped angle brackets in: %s", result)
	}
	if !strings.Contains(result, "&amp;") {
		t.Errorf("expected escaped ampersand in: %s", result)
	}
	if !strings.Contains(result, "&quot;chars&quot;") {
		t.Errorf("expected escaped quotes in: %s", result)
	}
}

func TestGolden_GenerateTOCXML(t *testing.T) {
	opts := TOCOptions{
		Title:         "Contents",
		OutlineLevels: "1-3",
	}

	result := string(generateTOCXML(opts))

	// Title paragraph must come before the field paragraph
	titleIdx := strings.Index(result, "Contents")
	fieldIdx := strings.Index(result, `w:fldCharType="begin"`)
	if titleIdx < 0 || fieldIdx < 0 {
		t.Fatalf("missing title or field in: %s", result)
	}
	if titleIdx >= fieldIdx {
		t.Error("title must appear before field begin")
	}

	// Verify TOC field instruction
	if !strings.Contains(result, `TOC \o &quot;1-3&quot;`) {
		t.Errorf("expected TOC field instruction in: %s", result)
	}

	// Verify field structure: begin, separate, end
	if !strings.Contains(result, `fldCharType="begin"`) {
		t.Error("missing fldChar begin")
	}
	if !strings.Contains(result, `fldCharType="separate"`) {
		t.Error("missing fldChar separate")
	}
	if !strings.Contains(result, `fldCharType="end"`) {
		t.Error("missing fldChar end")
	}
}

func TestGolden_GenerateTOCXML_NoTitle(t *testing.T) {
	opts := TOCOptions{
		OutlineLevels: "1-2",
	}

	result := string(generateTOCXML(opts))

	// Should not have a title paragraph
	if strings.Contains(result, "TOCHeading") {
		t.Error("should not have TOCHeading when title is empty")
	}
	// Should still have field
	if !strings.Contains(result, `fldCharType="begin"`) {
		t.Error("missing field begin")
	}
}

func TestGolden_GenerateStyleXML_Paragraph(t *testing.T) {
	def := StyleDefinition{
		ID:         "MyStyle",
		Name:       "My Custom Style",
		Type:       StyleTypeParagraph,
		BasedOn:    "Normal",
		NextStyle:  "Normal",
		FontFamily: "Arial",
		FontSize:   24,
		Color:      "FF0000",
		Bold:       true,
		Alignment:  ParagraphAlignCenter,
	}

	result := string(generateStyleXML(def))

	assertContains(t, result, `w:type="paragraph"`)
	assertContains(t, result, `w:styleId="MyStyle"`)
	assertContains(t, result, `<w:name w:val="My Custom Style"/>`)
	assertContains(t, result, `<w:basedOn w:val="Normal"/>`)
	assertContains(t, result, `<w:next w:val="Normal"/>`)
	assertContains(t, result, `<w:jc w:val="center"/>`)
	assertContains(t, result, `<w:rFonts w:ascii="Arial"`)
	assertContains(t, result, `<w:sz w:val="24"/>`)
	assertContains(t, result, `<w:color w:val="FF0000"/>`)
	assertContains(t, result, "<w:b/>")
}

func TestGolden_GenerateStyleXML_Character(t *testing.T) {
	def := StyleDefinition{
		ID:     "Highlight",
		Name:   "Highlight",
		Type:   StyleTypeCharacter,
		Bold:   true,
		Italic: true,
		Color:  "00FF00",
	}

	result := string(generateStyleXML(def))

	assertContains(t, result, `w:type="character"`)
	assertContains(t, result, "<w:b/>")
	assertContains(t, result, "<w:i/>")
	// Character styles should NOT have paragraph properties
	if strings.Contains(result, "<w:pPr>") {
		t.Error("character style should not have paragraph properties")
	}
}

func TestGolden_GenerateWatermarkXML(t *testing.T) {
	opts := WatermarkOptions{
		Text:       "CONFIDENTIAL",
		FontFamily: "Arial",
		Color:      "FF0000",
		Opacity:    0.5,
		Diagonal:   true,
	}

	result := string(generateWatermarkShapeXML(opts))

	assertContains(t, result, "CONFIDENTIAL")
	assertContains(t, result, "rotation:315")
	assertContains(t, result, `fillcolor="#FF0000"`)
	assertContains(t, result, `opacity="0.50"`)
	assertContains(t, result, `font-family`)
	assertContains(t, result, "<v:shapetype")
	assertContains(t, result, "<v:shape")
}

func TestGolden_GenerateWatermarkXML_NoDiagonal(t *testing.T) {
	opts := WatermarkOptions{
		Text:    "DRAFT",
		Color:   "C0C0C0",
		Opacity: 0.3,
	}

	result := string(generateWatermarkShapeXML(opts))

	if strings.Contains(result, "rotation:315") {
		t.Error("should not have rotation when Diagonal is false")
	}
	assertContains(t, result, "DRAFT")
}

func TestGolden_GenerateCommentEntry(t *testing.T) {
	opts := CommentOptions{
		Text:     "Please review this section",
		Author:   "Reviewer",
		Initials: "R",
	}

	result := string(generateCommentEntry(42, opts))

	assertContains(t, result, `w:id="42"`)
	assertContains(t, result, `w:author="Reviewer"`)
	assertContains(t, result, `w:initials="R"`)
	assertContains(t, result, "Please review this section")
	assertContains(t, result, `<w:pStyle w:val="CommentText"/>`)
	assertContains(t, result, "<w:annotationRef/>")
}

func TestGolden_GenerateFootnoteEntry(t *testing.T) {
	result := string(generateFootnoteEntry(7, "See appendix A for details."))

	assertContains(t, result, `w:id="7"`)
	assertContains(t, result, `<w:pStyle w:val="FootnoteText"/>`)
	assertContains(t, result, "<w:footnoteRef/>")
	assertContains(t, result, "See appendix A for details.")
}

func TestGolden_GenerateEndnoteEntry(t *testing.T) {
	result := string(generateEndnoteEntry(3, "Full reference available online."))

	assertContains(t, result, `w:id="3"`)
	assertContains(t, result, `<w:pStyle w:val="EndnoteText"/>`)
	assertContains(t, result, "<w:endnoteRef/>")
	assertContains(t, result, "Full reference available online.")
}

func TestGolden_GenerateTrackedInsertXML(t *testing.T) {
	opts := TrackedInsertOptions{
		Text:   "New content",
		Author: "Editor",
		Date:   time.Date(2026, 3, 15, 12, 0, 0, 0, time.UTC),
		Style:  StyleHeading2,
		Bold:   true,
	}

	result := string(generateTrackedInsertXML(opts))

	assertContains(t, result, `<w:ins w:id="2" w:author="Editor"`)
	assertContains(t, result, "2026-03-15T12:00:00Z")
	assertContains(t, result, `<w:pStyle w:val="Heading2"/>`)
	assertContains(t, result, "<w:b/>")
	assertContains(t, result, "New content")
	assertContains(t, result, "</w:ins>")
}

func TestGolden_SetPageNumberInSectPr(t *testing.T) {
	input := []byte(`<w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>`)
	opts := PageNumberOptions{Start: 5, Format: PageNumUpperRoman}

	result, err := setPageNumberInSectPr(input, opts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	rs := string(result)
	assertContains(t, rs, `w:start="5"`)
	assertContains(t, rs, `w:fmt="upperRoman"`)
}

func TestGolden_InjectTcPrElement(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		element  string
		contains string
	}{
		{
			name:     "no existing tcPr",
			input:    "<w:tc><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:tc>",
			element:  `<w:gridSpan w:val="3"/>`,
			contains: `<w:tcPr><w:gridSpan w:val="3"/></w:tcPr>`,
		},
		{
			name:     "existing tcPr",
			input:    "<w:tc><w:tcPr><w:tcW w:w=\"1000\"/></w:tcPr><w:p/></w:tc>",
			element:  `<w:vMerge w:val="restart"/>`,
			contains: `<w:vMerge w:val="restart"/>`,
		},
		{
			name:     "self-closing tcPr",
			input:    "<w:tc><w:tcPr/><w:p/></w:tc>",
			element:  `<w:gridSpan w:val="2"/>`,
			contains: `<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := injectTcPrElement(tt.input, tt.element)
			if !strings.Contains(result, tt.contains) {
				t.Errorf("expected %q in: %s", tt.contains, result)
			}
		})
	}
}

func TestGolden_MarkTOCForUpdate(t *testing.T) {
	input := []byte(`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
		`<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r>` +
		`<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>`)

	result := markTOCForUpdate(input)

	rs := string(result)
	assertContains(t, rs, `w:dirty="true"`)
}

func TestGolden_EnsureVMLNamespaces(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>`

	result := ensureVMLNamespaces(input)

	assertContains(t, result, `xmlns:v="urn:schemas-microsoft-com:vml"`)
	assertContains(t, result, `xmlns:o="urn:schemas-microsoft-com:office:office"`)
}

func TestGolden_XMLEscape(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"hello", "hello"},
		{"a & b", "a &amp; b"},
		{"<tag>", "&lt;tag&gt;"},
		{`say "hi"`, `say &quot;hi&quot;`},
		{"it's", "it&apos;s"},
		{"a < b & c > d", "a &lt; b &amp; c &gt; d"},
	}

	for _, tt := range tests {
		result := xmlEscape(tt.input)
		if result != tt.expected {
			t.Errorf("xmlEscape(%q) = %q, want %q", tt.input, result, tt.expected)
		}
	}
}

func TestGolden_NormalizeHexColor(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"FF0000", "FF0000"},
		{"ff0000", "FF0000"},
		{"#FF0000", "FF0000"},
		{"#ff0000", "FF0000"},
		{"invalid", ""},
		{"FFFF", ""},
		{"", ""},
	}

	for _, tt := range tests {
		result := normalizeHexColor(tt.input)
		if result != tt.expected {
			t.Errorf("normalizeHexColor(%q) = %q, want %q", tt.input, result, tt.expected)
		}
	}
}

func TestGolden_ColumnLetters(t *testing.T) {
	tests := []struct {
		input    int
		expected string
	}{
		{1, "A"},
		{2, "B"},
		{26, "Z"},
		{27, "AA"},
		{28, "AB"},
	}

	for _, tt := range tests {
		result := columnLetters(tt.input)
		if result != tt.expected {
			t.Errorf("columnLetters(%d) = %q, want %q", tt.input, result, tt.expected)
		}
	}
}

// assertContains is a test helper that checks if s contains substr.
func assertContains(t *testing.T, s, substr string) {
	t.Helper()
	if !strings.Contains(s, substr) {
		t.Errorf("expected %q in output:\n%s", substr, s)
	}
}

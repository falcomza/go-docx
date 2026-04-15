package godocx

import (
	"bytes"
	"strings"
	"testing"
)

func TestGenerateTOCXML_DefaultOptions(t *testing.T) {
	opts := DefaultTOCOptions()
	result := generateTOCXML(opts)
	xml := string(result)

	// Title should come before the TOC field
	titleIdx := strings.Index(xml, "Table of Contents")
	fieldIdx := strings.Index(xml, `fldCharType="begin"`)
	if titleIdx == -1 {
		t.Error("expected title paragraph in TOC XML")
	}
	if fieldIdx == -1 {
		t.Error("expected field begin in TOC XML")
	}
	if titleIdx > fieldIdx {
		t.Error("expected title to come before field")
	}

	// Check field instruction contains proper TOC switches
	if !strings.Contains(xml, `TOC \o`) {
		t.Error("expected TOC field instruction with \\o switch")
	}
	if !strings.Contains(xml, `\h`) {
		t.Error("expected TOC field instruction with \\h switch")
	}
	if !strings.Contains(xml, `\z`) {
		t.Error("expected TOC field instruction with \\z switch")
	}

	// Check field structure: begin, instrText, separate, result, end
	if !strings.Contains(xml, `fldCharType="begin"`) {
		t.Error("expected fldChar begin")
	}
	if !strings.Contains(xml, `fldCharType="separate"`) {
		t.Error("expected fldChar separate")
	}
	if !strings.Contains(xml, `fldCharType="end"`) {
		t.Error("expected fldChar end")
	}
}

func TestGenerateTOCXML_NoTitle(t *testing.T) {
	opts := TOCOptions{
		OutlineLevels: "1-2",
	}
	result := generateTOCXML(opts)
	xml := string(result)

	// Should not contain TOCHeading style since no title
	if strings.Contains(xml, "TOCHeading") {
		t.Error("expected no TOCHeading style when title is empty")
	}

	// Should still have the field
	if !strings.Contains(xml, `fldCharType="begin"`) {
		t.Error("expected fldChar begin")
	}

	// Check outline levels are correct
	if !strings.Contains(xml, `1-2`) {
		t.Error("expected outline levels 1-2 in field instruction")
	}
}

func TestGenerateTOCXML_CustomOutlineLevels(t *testing.T) {
	opts := TOCOptions{
		OutlineLevels: "1-5",
	}
	result := generateTOCXML(opts)
	xml := string(result)

	if !strings.Contains(xml, "1-5") {
		t.Error("expected outline levels 1-5 in field instruction")
	}
}

func TestGenerateCaptionListXML_TableOfFigures(t *testing.T) {
	opts := DefaultTableOfFiguresOptions()
	result := generateCaptionListXML(opts, CaptionFigure)
	xml := string(result)

	if !strings.Contains(xml, "Table of Figures") {
		t.Error("expected title paragraph for table of figures")
	}
	if !strings.Contains(xml, `TOC \h \z \c &quot;Figure&quot;`) && !strings.Contains(xml, `TOC \h \z \c "Figure"`) {
		t.Error("expected TOC field instruction with figure caption switch")
	}
	if !strings.Contains(xml, `fldCharType="begin"`) || !strings.Contains(xml, `fldCharType="end"`) {
		t.Error("expected field begin/end in caption list XML")
	}
}

func TestGenerateCaptionListXML_TableOfTables(t *testing.T) {
	opts := DefaultTableOfTablesOptions()
	result := generateCaptionListXML(opts, CaptionTable)
	xml := string(result)

	if !strings.Contains(xml, "Table of Tables") {
		t.Error("expected title paragraph for table of tables")
	}
	if !strings.Contains(xml, `TOC \h \z \c &quot;Table&quot;`) && !strings.Contains(xml, `TOC \h \z \c "Table"`) {
		t.Error("expected TOC field instruction with table caption switch")
	}
}

func TestMarkTOCForUpdate(t *testing.T) {
	// Create a document XML with a TOC field
	docXML := []byte(`<w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
		`<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z </w:instrText></w:r>` +
		`<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
		`<w:r><w:t>placeholder</w:t></w:r>` +
		`<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body>`)

	result := markTOCForUpdate(docXML)

	if !bytes.Contains(result, []byte(`w:dirty="true"`)) {
		t.Error("expected dirty attribute to be added")
	}
	if !bytes.Contains(result, []byte(`fldCharType="begin" w:dirty="true"`)) {
		t.Error("expected dirty attribute on begin fldChar")
	}
}

func TestMarkTOCForUpdate_NoTOC(t *testing.T) {
	docXML := []byte(`<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>`)

	result := markTOCForUpdate(docXML)

	if !bytes.Equal(result, docXML) {
		t.Error("expected no changes when no TOC field")
	}
}

func TestMarkTOCForUpdate_AlreadyDirty(t *testing.T) {
	docXML := []byte(`<w:body><w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>` +
		`<w:r><w:instrText> TOC \o "1-3" </w:instrText></w:r></w:p></w:body>`)

	result := markTOCForUpdate(docXML)

	// Should not add another dirty attribute
	count := bytes.Count(result, []byte(`w:dirty`))
	if count != 1 {
		t.Errorf("expected 1 dirty attribute, got %d", count)
	}
}

func TestParseTOCEntries(t *testing.T) {
	docXML := []byte(`<w:body>` +
		`<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Chapter 1</w:t></w:r></w:p>` +
		`<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Section 1.1</w:t></w:r></w:p>` +
		`<w:p><w:pPr><w:pStyle w:val="TOC3"/></w:pPr><w:r><w:t>Subsection 1.1.1</w:t></w:r></w:p>` +
		`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Regular text</w:t></w:r></w:p>` +
		`</w:body>`)

	entries := parseTOCEntries(docXML)

	if len(entries) != 3 {
		t.Fatalf("expected 3 TOC entries, got %d", len(entries))
	}

	if entries[0].Level != 1 || entries[0].Text != "Chapter 1" {
		t.Errorf("unexpected entry 0: %+v", entries[0])
	}
	if entries[1].Level != 2 || entries[1].Text != "Section 1.1" {
		t.Errorf("unexpected entry 1: %+v", entries[1])
	}
	if entries[2].Level != 3 || entries[2].Text != "Subsection 1.1.1" {
		t.Errorf("unexpected entry 2: %+v", entries[2])
	}
}

func TestGenerateWatermarkShapeXML(t *testing.T) {
	opts := DefaultWatermarkOptions()
	result := generateWatermarkShapeXML(opts)
	xml := string(result)

	if !strings.Contains(xml, "DRAFT") {
		t.Error("expected watermark text DRAFT")
	}
	if !strings.Contains(xml, "v:shape") {
		t.Error("expected VML shape element")
	}
	if !strings.Contains(xml, "v:shapetype") {
		t.Error("expected VML shapetype definition")
	}
	if !strings.Contains(xml, "rotation:315") {
		t.Error("expected diagonal rotation")
	}
	if !strings.Contains(xml, "Calibri") {
		t.Error("expected default font family")
	}
}

func TestGenerateWatermarkShapeXML_NoDiagonal(t *testing.T) {
	opts := WatermarkOptions{
		Text:       "CONFIDENTIAL",
		FontFamily: "Arial",
		Color:      "FF0000",
		Opacity:    0.3,
		Diagonal:   false,
	}
	result := generateWatermarkShapeXML(opts)
	xml := string(result)

	if !strings.Contains(xml, "CONFIDENTIAL") {
		t.Error("expected watermark text CONFIDENTIAL")
	}
	if strings.Contains(xml, "rotation") {
		t.Error("expected no rotation when Diagonal is false")
	}
	if !strings.Contains(xml, "FF0000") {
		t.Error("expected red color")
	}
}

func TestSetPageNumberInSectPr_NewSectPr(t *testing.T) {
	docXML := []byte(`<?xml version="1.0"?><w:body><w:p/></w:body>`)
	opts := PageNumberOptions{Start: 1, Format: PageNumDecimal}

	result, err := setPageNumberInSectPr(docXML, opts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if !bytes.Contains(result, []byte(`<w:pgNumType w:start="1" w:fmt="decimal"/>`)) {
		t.Error("expected pgNumType element in result")
	}
	if !bytes.Contains(result, []byte(`<w:sectPr>`)) {
		t.Error("expected new sectPr to be created")
	}
}

func TestSetPageNumberInSectPr_ExistingSectPr(t *testing.T) {
	docXML := []byte(`<?xml version="1.0"?><w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>`)
	opts := PageNumberOptions{Start: 5, Format: PageNumUpperRoman}

	result, err := setPageNumberInSectPr(docXML, opts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if !bytes.Contains(result, []byte(`<w:pgNumType w:start="5" w:fmt="upperRoman"/>`)) {
		t.Error("expected pgNumType with start=5 and upperRoman format")
	}
}

func TestSetPageNumberInSectPr_ReplacesExisting(t *testing.T) {
	docXML := []byte(`<?xml version="1.0"?><w:body><w:p/><w:sectPr><w:pgNumType w:start="1" w:fmt="decimal"/></w:sectPr></w:body>`)
	opts := PageNumberOptions{Start: 10, Format: PageNumLowerRoman}

	result, err := setPageNumberInSectPr(docXML, opts)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if !bytes.Contains(result, []byte(`<w:pgNumType w:start="10" w:fmt="lowerRoman"/>`)) {
		t.Error("expected pgNumType to be replaced")
	}
	// Ensure old one is gone
	if bytes.Contains(result, []byte(`w:start="1"`)) {
		t.Error("old pgNumType should have been replaced")
	}
}

func TestEnsureVMLNamespaces(t *testing.T) {
	headerXML := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>`
	result := ensureVMLNamespaces(headerXML)

	if !strings.Contains(result, `xmlns:v="urn:schemas-microsoft-com:vml"`) {
		t.Error("expected VML namespace to be added")
	}
	if !strings.Contains(result, `xmlns:o="urn:schemas-microsoft-com:office:office"`) {
		t.Error("expected Office namespace to be added")
	}
}

func TestEnsureVMLNamespaces_AlreadyPresent(t *testing.T) {
	headerXML := `<w:hdr xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>`
	result := ensureVMLNamespaces(headerXML)

	// Should not add duplicate
	count := strings.Count(result, `xmlns:v=`)
	if count != 1 {
		t.Errorf("expected 1 xmlns:v declaration, got %d", count)
	}
}

package godocx

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// --- Fixture helpers ---

// buildIntegrationFixture creates a minimal DOCX as bytes with the given document.xml body content.
func buildIntegrationFixture(t *testing.T, bodyContent string) []byte {
	t.Helper()

	docXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
		`xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
		`xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ` +
		`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
		`xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
		`<w:body>` + bodyContent + `</w:body></w:document>`

	return buildIntegrationDocxFromParts(t, docXML, "", "")
}

// buildIntegrationDocxFromParts creates a DOCX zip from doc XML and optional extra files.
func buildIntegrationDocxFromParts(t *testing.T, docXML, stylesXML, contentTypesOverride string) []byte {
	t.Helper()

	buf := &bytes.Buffer{}
	w := zip.NewWriter(buf)

	ct := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`
	if contentTypesOverride != "" {
		ct = contentTypesOverride
	}
	writeZipStr(t, w, "[Content_Types].xml", ct)
	writeZipStr(t, w, "word/document.xml", docXML)
	writeZipStr(t, w, "word/_rels/document.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
			`</Relationships>`)

	if stylesXML != "" {
		writeZipStr(t, w, "word/styles.xml", stylesXML)
	}

	if err := w.Close(); err != nil {
		t.Fatalf("close zip: %v", err)
	}
	return buf.Bytes()
}

func writeZipStr(t *testing.T, w *zip.Writer, name, content string) {
	t.Helper()
	f, err := w.Create(name)
	if err != nil {
		t.Fatalf("create zip entry %s: %v", name, err)
	}
	if _, err := f.Write([]byte(content)); err != nil {
		t.Fatalf("write zip entry %s: %v", name, err)
	}
}

// newUpdaterFromFixture writes a fixture to disk and opens it with New().
func newUpdaterFromFixture(t *testing.T, fixture []byte) *Updater {
	t.Helper()
	path := filepath.Join(t.TempDir(), "input.docx")
	if err := os.WriteFile(path, fixture, 0o644); err != nil {
		t.Fatalf("write fixture: %v", err)
	}
	u, err := New(path)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	t.Cleanup(func() { u.Cleanup() })
	return u
}

// readDocXML reads the document.xml from the updater's tempDir.
func readDocXML(t *testing.T, u *Updater) string {
	t.Helper()
	raw, err := os.ReadFile(filepath.Join(u.TempDir(), "word", "document.xml"))
	if err != nil {
		t.Fatalf("read document.xml: %v", err)
	}
	return string(raw)
}

// --- Delete operations integration tests ---

func TestDeleteParagraphs_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Keep this</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Remove me</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Also keep</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	count, err := u.DeleteParagraphs("Remove me", DefaultDeleteOptions())
	if err != nil {
		t.Fatalf("DeleteParagraphs: %v", err)
	}
	if count != 1 {
		t.Errorf("expected 1 deletion, got %d", count)
	}

	docXML := readDocXML(t, u)
	if strings.Contains(docXML, "Remove me") {
		t.Error("deleted paragraph still present")
	}
	if !strings.Contains(docXML, "Keep this") {
		t.Error("kept paragraph missing")
	}
	if !strings.Contains(docXML, "Also keep") {
		t.Error("second kept paragraph missing")
	}
}

func TestDeleteParagraphs_EmptyText(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>text</w:t></w:r></w:p>`))
	_, err := u.DeleteParagraphs("", DefaultDeleteOptions())
	if err == nil {
		t.Error("expected error for empty text")
	}
}

func TestDeleteParagraphs_CaseSensitive(t *testing.T) {
	body := `<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>hello world</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	count, err := u.DeleteParagraphs("Hello", DeleteOptions{MatchCase: true})
	if err != nil {
		t.Fatalf("DeleteParagraphs: %v", err)
	}
	if count != 1 {
		t.Errorf("expected 1 deletion (case-sensitive), got %d", count)
	}
}

func TestDeleteParagraphs_WholeWord(t *testing.T) {
	body := `<w:p><w:r><w:t>The cat sat</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>category list</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	count, err := u.DeleteParagraphs("cat", DeleteOptions{WholeWord: true})
	if err != nil {
		t.Fatalf("DeleteParagraphs: %v", err)
	}
	if count != 1 {
		t.Errorf("expected 1 deletion (whole word), got %d", count)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "category") {
		t.Error("'category' should not be matched with whole word")
	}
}

func TestDeleteTable_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Before table</w:t></w:r></w:p>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:p><w:r><w:t>After table</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.DeleteTable(1)
	if err != nil {
		t.Fatalf("DeleteTable: %v", err)
	}

	docXML := readDocXML(t, u)
	if strings.Contains(docXML, "<w:tbl>") {
		t.Error("table still present after deletion")
	}
	if !strings.Contains(docXML, "Before table") || !strings.Contains(docXML, "After table") {
		t.Error("surrounding text should be preserved")
	}
}

func TestDeleteTable_NotFound(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>No table</w:t></w:r></w:p>`))
	err := u.DeleteTable(1)
	if err == nil {
		t.Error("expected error when deleting non-existent table")
	}
}

func TestDeleteImage_Integration(t *testing.T) {
	body := `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
		`<wp:extent cx="100" cy="100"/>` +
		`<wp:docPr id="1" name="Pic1"/>` +
		`<a:graphic><a:graphicData><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
		`</wp:inline></w:drawing></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.DeleteImage(1)
	if err != nil {
		t.Fatalf("DeleteImage: %v", err)
	}

	docXML := readDocXML(t, u)
	if strings.Contains(docXML, "r:embed") {
		t.Error("image still present after deletion")
	}
}

func TestDeleteImage_NotFound(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>No image</w:t></w:r></w:p>`))
	err := u.DeleteImage(1)
	if err == nil {
		t.Error("expected error when deleting non-existent image")
	}
}

func TestDeleteChart_Integration(t *testing.T) {
	body := `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
		`<wp:extent cx="100" cy="100"/>` +
		`<wp:docPr id="1" name="Chart1"/>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
		`<c:chart r:id="rId10"/></a:graphicData></a:graphic>` +
		`</wp:inline></w:drawing></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.DeleteChart(1)
	if err != nil {
		t.Fatalf("DeleteChart: %v", err)
	}

	docXML := readDocXML(t, u)
	if strings.Contains(docXML, "c:chart") {
		t.Error("chart still present after deletion")
	}
}

func TestDeleteChart_NotFound(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>No chart</w:t></w:r></w:p>`))
	err := u.DeleteChart(1)
	if err == nil {
		t.Error("expected error when deleting non-existent chart")
	}
}

// --- Count operations integration tests ---

func TestGetTableCount_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>text</w:t></w:r></w:p>` +
		`<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
		`<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	count, err := u.GetTableCount()
	if err != nil {
		t.Fatalf("GetTableCount: %v", err)
	}
	if count != 2 {
		t.Errorf("expected 2 tables, got %d", count)
	}
}

func TestGetParagraphCount_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Para 1</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Para 2</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Para 3</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	count, err := u.GetParagraphCount()
	if err != nil {
		t.Fatalf("GetParagraphCount: %v", err)
	}
	if count != 3 {
		t.Errorf("expected 3 paragraphs, got %d", count)
	}
}

func TestGetImageCount_Integration(t *testing.T) {
	body := `<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData>` +
		`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>` +
		`</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	count, err := u.GetImageCount()
	if err != nil {
		t.Fatalf("GetImageCount: %v", err)
	}
	if count != 1 {
		t.Errorf("expected 1 image, got %d", count)
	}
}

// --- TOC integration tests ---

func TestInsertTOC_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Introduction</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTOC(TOCOptions{
		Title:         "Contents",
		OutlineLevels: "1-3",
		Position:      PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertTOC: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "Contents") {
		t.Error("TOC title not found")
	}
	if !strings.Contains(docXML, "fldCharType") {
		t.Error("TOC field not found")
	}
	if !strings.Contains(docXML, "TOC") {
		t.Error("TOC instruction not found")
	}
}

func TestInsertTOC_AtEnd(t *testing.T) {
	body := `<w:p><w:r><w:t>Introduction</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTOC(TOCOptions{
		Title:    "Table of Contents",
		Position: PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertTOC at end: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "Table of Contents") {
		t.Error("TOC title not found")
	}
}

func TestInsertTOC_AfterAnchor(t *testing.T) {
	body := `<w:p><w:r><w:t>Chapter 1</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Chapter 2</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTOC(TOCOptions{
		Title:    "",
		Position: PositionAfterText,
		Anchor:   "Chapter 1",
	})
	if err != nil {
		t.Fatalf("InsertTOC after anchor: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "fldCharType") {
		t.Error("TOC field not found")
	}
}

func TestUpdateTOC_Integration(t *testing.T) {
	// Create a document that already has a TOC field
	tocBody := `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
		`<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r>` +
		`<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
		`<w:r><w:t>Placeholder</w:t></w:r>` +
		`<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, tocBody))

	err := u.UpdateTOC()
	if err != nil {
		t.Fatalf("UpdateTOC: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, `w:dirty="true"`) {
		t.Error("TOC not marked as dirty")
	}
}

func TestGetTOCEntries_Integration(t *testing.T) {
	body := `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Chapter One</w:t></w:r></w:p>` +
		`<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Section A</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Regular paragraph</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	entries, err := u.GetTOCEntries()
	if err != nil {
		t.Fatalf("GetTOCEntries: %v", err)
	}
	if len(entries) != 2 {
		t.Fatalf("expected 2 TOC entries, got %d", len(entries))
	}
	if entries[0].Level != 1 || !strings.Contains(entries[0].Text, "Chapter One") {
		t.Errorf("first entry: level=%d text=%q", entries[0].Level, entries[0].Text)
	}
	if entries[1].Level != 2 || !strings.Contains(entries[1].Text, "Section A") {
		t.Errorf("second entry: level=%d text=%q", entries[1].Level, entries[1].Text)
	}
}

func TestInsertTableOfFigures_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Introduction</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTableOfFigures(DefaultTableOfFiguresOptions())
	if err != nil {
		t.Fatalf("InsertTableOfFigures: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "Table of Figures") {
		t.Error("table of figures title not found")
	}
	if !strings.Contains(docXML, `TOC \h \z \c &quot;Figure&quot;`) && !strings.Contains(docXML, `TOC \h \z \c "Figure"`) {
		t.Error("figure caption list instruction not found")
	}
	if !strings.Contains(docXML, "fldCharType") {
		t.Error("field code not found for table of figures")
	}
}

func TestInsertTableOfTables_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Introduction</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTableOfTables(DefaultTableOfTablesOptions())
	if err != nil {
		t.Fatalf("InsertTableOfTables: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "Table of Tables") {
		t.Error("table of tables title not found")
	}
	if !strings.Contains(docXML, `TOC \h \z \c &quot;Table&quot;`) && !strings.Contains(docXML, `TOC \h \z \c "Table"`) {
		t.Error("table caption list instruction not found")
	}
	if !strings.Contains(docXML, "fldCharType") {
		t.Error("field code not found for table of tables")
	}
}

func TestGenerateDocxWithAllThreeTables(t *testing.T) {
	body := `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter 1</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Body text</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	if err := u.InsertTOC(TOCOptions{
		Title:         "Table of Contents",
		OutlineLevels: "1-3",
		Position:      PositionBeginning,
	}); err != nil {
		t.Fatalf("InsertTOC: %v", err)
	}

	figOpts := DefaultTableOfFiguresOptions()
	figOpts.Position = PositionEnd
	if err := u.InsertTableOfFigures(figOpts); err != nil {
		t.Fatalf("InsertTableOfFigures: %v", err)
	}

	tableOpts := DefaultTableOfTablesOptions()
	tableOpts.Position = PositionEnd
	if err := u.InsertTableOfTables(tableOpts); err != nil {
		t.Fatalf("InsertTableOfTables: %v", err)
	}

	outputPath := filepath.Join(t.TempDir(), "all_three_tables.docx")
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	zr, err := zip.OpenReader(outputPath)
	if err != nil {
		t.Fatalf("open saved docx: %v", err)
	}
	defer zr.Close()

	var docXML string
	for _, f := range zr.File {
		if f.Name != "word/document.xml" {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open document.xml in zip: %v", err)
		}
		raw, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read document.xml from zip: %v", err)
		}
		docXML = string(raw)
		break
	}

	if docXML == "" {
		t.Fatal("document.xml not found in saved docx")
	}

	if !strings.Contains(docXML, "Table of Contents") {
		t.Error("missing Table of Contents title")
	}
	if !strings.Contains(docXML, "Table of Figures") {
		t.Error("missing Table of Figures title")
	}
	if !strings.Contains(docXML, "Table of Tables") {
		t.Error("missing Table of Tables title")
	}

	if !strings.Contains(docXML, `TOC \o &quot;1-3&quot;`) && !strings.Contains(docXML, `TOC \o "1-3"`) {
		t.Error("missing heading TOC field instruction")
	}
	if !strings.Contains(docXML, `TOC \h \z \c &quot;Figure&quot;`) && !strings.Contains(docXML, `TOC \h \z \c "Figure"`) {
		t.Error("missing table of figures field instruction")
	}
	if !strings.Contains(docXML, `TOC \h \z \c &quot;Table&quot;`) && !strings.Contains(docXML, `TOC \h \z \c "Table"`) {
		t.Error("missing table of tables field instruction")
	}

	if got := strings.Count(docXML, `w:fldCharType="begin"`); got < 3 {
		t.Errorf("expected at least 3 field begin markers, got %d", got)
	}
}

// --- Watermark integration tests ---

func TestSetTextWatermark_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.SetTextWatermark(WatermarkOptions{
		Text:     "DRAFT",
		Color:    "C0C0C0",
		Opacity:  0.5,
		Diagonal: true,
	})
	if err != nil {
		t.Fatalf("SetTextWatermark: %v", err)
	}

	// Verify header file was created
	headerFound := false
	wordDir := filepath.Join(u.TempDir(), "word")
	entries, _ := os.ReadDir(wordDir)
	for _, e := range entries {
		if strings.HasPrefix(e.Name(), "header") && strings.HasSuffix(e.Name(), ".xml") {
			headerPath := filepath.Join(wordDir, e.Name())
			raw, err := os.ReadFile(headerPath)
			if err == nil && strings.Contains(string(raw), "DRAFT") {
				headerFound = true
			}
		}
	}
	if !headerFound {
		t.Error("watermark header not found")
	}

	// Verify sectPr updated with header reference
	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "headerReference") {
		t.Error("header reference not found in document")
	}
}

func TestSetTextWatermark_EmptyText(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))
	err := u.SetTextWatermark(WatermarkOptions{Text: ""})
	if err == nil {
		t.Error("expected error for empty watermark text")
	}
}

func TestSetTextWatermark_WithExistingHeader(t *testing.T) {
	// Create fixture with an existing default header
	buf := &bytes.Buffer{}
	w := zip.NewWriter(buf)

	docXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<w:body><w:p><w:r><w:t>Content</w:t></w:r></w:p>` +
		`<w:sectPr><w:headerReference w:type="default" r:id="rId1"/>` +
		`<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`

	headerXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<w:p><w:r><w:t>Existing Header</w:t></w:r></w:p></w:hdr>`

	relsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
		`</Relationships>`

	writeZipStr(t, w, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	writeZipStr(t, w, "word/document.xml", docXML)
	writeZipStr(t, w, "word/_rels/document.xml.rels", relsXML)
	writeZipStr(t, w, "word/header1.xml", headerXML)

	if err := w.Close(); err != nil {
		t.Fatalf("close zip: %v", err)
	}

	u := newUpdaterFromFixture(t, buf.Bytes())

	err := u.SetTextWatermark(WatermarkOptions{
		Text:     "CONFIDENTIAL",
		Color:    "FF0000",
		Opacity:  0.7,
		Diagonal: false,
	})
	if err != nil {
		t.Fatalf("SetTextWatermark with existing header: %v", err)
	}

	// Verify watermark was injected into existing header
	headerRaw, err := os.ReadFile(filepath.Join(u.TempDir(), "word", "header1.xml"))
	if err != nil {
		t.Fatalf("read header: %v", err)
	}
	headerContent := string(headerRaw)
	if !strings.Contains(headerContent, "CONFIDENTIAL") {
		t.Error("watermark not injected into existing header")
	}
	if !strings.Contains(headerContent, "Existing Header") {
		t.Error("existing header content lost")
	}
}

// --- Page number integration tests ---

func TestSetPageNumber_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.SetPageNumber(PageNumberOptions{
		Start:  5,
		Format: PageNumUpperRoman,
	})
	if err != nil {
		t.Fatalf("SetPageNumber: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, `w:start="5"`) {
		t.Error("page number start not set")
	}
	if !strings.Contains(docXML, `w:fmt="upperRoman"`) {
		t.Error("page number format not set")
	}
}

func TestSetPageNumber_NoSectPr(t *testing.T) {
	// Document without sectPr - should create one
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.SetPageNumber(PageNumberOptions{
		Start:  1,
		Format: PageNumDecimal,
	})
	if err != nil {
		t.Fatalf("SetPageNumber: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "w:pgNumType") {
		t.Error("pgNumType not created")
	}
}

func TestSetPageNumber_NegativeStart(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))
	err := u.SetPageNumber(PageNumberOptions{Start: -1})
	if err == nil {
		t.Error("expected error for negative page start")
	}
}

// --- Footnote/Endnote integration tests ---

func TestInsertFootnote_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>See reference here</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertFootnote(FootnoteOptions{
		Text:   "This is footnote text.",
		Anchor: "reference here",
	})
	if err != nil {
		t.Fatalf("InsertFootnote: %v", err)
	}

	// Verify footnotes.xml created
	fnPath := filepath.Join(u.TempDir(), "word", "footnotes.xml")
	raw, err := os.ReadFile(fnPath)
	if err != nil {
		t.Fatalf("read footnotes.xml: %v", err)
	}
	if !strings.Contains(string(raw), "This is footnote text.") {
		t.Error("footnote text not found")
	}

	// Verify footnote reference in document
	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "footnoteReference") {
		t.Error("footnote reference not found in document")
	}
}

func TestInsertFootnote_EmptyText(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>text</w:t></w:r></w:p>`))
	err := u.InsertFootnote(FootnoteOptions{Text: "", Anchor: "text"})
	if err == nil {
		t.Error("expected error for empty footnote text")
	}
}

func TestInsertFootnote_EmptyAnchor(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>text</w:t></w:r></w:p>`))
	err := u.InsertFootnote(FootnoteOptions{Text: "note", Anchor: ""})
	if err == nil {
		t.Error("expected error for empty anchor")
	}
}

func TestInsertEndnote_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>See endnote here</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertEndnote(EndnoteOptions{
		Text:   "This is endnote text.",
		Anchor: "endnote here",
	})
	if err != nil {
		t.Fatalf("InsertEndnote: %v", err)
	}

	// Verify endnotes.xml created
	enPath := filepath.Join(u.TempDir(), "word", "endnotes.xml")
	raw, err := os.ReadFile(enPath)
	if err != nil {
		t.Fatalf("read endnotes.xml: %v", err)
	}
	if !strings.Contains(string(raw), "This is endnote text.") {
		t.Error("endnote text not found")
	}

	// Verify endnote reference in document
	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "endnoteReference") {
		t.Error("endnote reference not found in document")
	}
}

func TestInsertEndnote_EmptyText(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p><w:r><w:t>text</w:t></w:r></w:p>`))
	err := u.InsertEndnote(EndnoteOptions{Text: "", Anchor: "text"})
	if err == nil {
		t.Error("expected error for empty endnote text")
	}
}

// --- Style integration tests ---

func TestAddStyle_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.AddStyle(StyleDefinition{
		ID:        "CustomHeading",
		Name:      "Custom Heading",
		Type:      StyleTypeParagraph,
		BasedOn:   "Normal",
		Bold:      true,
		FontSize:  28,
		Color:     "1F4E79",
		Alignment: ParagraphAlignCenter,
		KeepNext:  true,
	})
	if err != nil {
		t.Fatalf("AddStyle: %v", err)
	}

	// Verify styles.xml was created
	stylesPath := filepath.Join(u.TempDir(), "word", "styles.xml")
	raw, err := os.ReadFile(stylesPath)
	if err != nil {
		t.Fatalf("read styles.xml: %v", err)
	}
	contents := string(raw)
	if !strings.Contains(contents, "CustomHeading") {
		t.Error("custom style not found in styles.xml")
	}
	if !strings.Contains(contents, "<w:b/>") {
		t.Error("bold formatting not found")
	}
}

func TestAddStyle_WithExistingStyles(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>`

	stylesXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
		`</w:styles>`

	fixture := buildIntegrationDocxFromParts(t,
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`+
			`<w:body>`+body+`</w:body></w:document>`,
		stylesXML, "")

	u := newUpdaterFromFixture(t, fixture)

	err := u.AddStyle(StyleDefinition{
		ID:   "Emphasis",
		Name: "Emphasis",
		Type: StyleTypeCharacter,
		Bold: true,
	})
	if err != nil {
		t.Fatalf("AddStyle with existing: %v", err)
	}

	stylesPath := filepath.Join(u.TempDir(), "word", "styles.xml")
	raw, _ := os.ReadFile(stylesPath)
	contents := string(raw)
	if !strings.Contains(contents, "Emphasis") {
		t.Error("new style not injected")
	}
	if !strings.Contains(contents, "Normal") {
		t.Error("existing style lost")
	}
}

func TestAddStyle_EmptyID(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))
	err := u.AddStyle(StyleDefinition{ID: ""})
	if err == nil {
		t.Error("expected error for empty style ID")
	}
}

func TestAddStyles_Batch(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))

	err := u.AddStyles([]StyleDefinition{
		{ID: "Style1", Bold: true},
		{ID: "Style2", Italic: true},
	})
	if err != nil {
		t.Fatalf("AddStyles: %v", err)
	}

	stylesPath := filepath.Join(u.TempDir(), "word", "styles.xml")
	raw, _ := os.ReadFile(stylesPath)
	contents := string(raw)
	if !strings.Contains(contents, "Style1") || !strings.Contains(contents, "Style2") {
		t.Error("batch styles not all added")
	}
}

// --- Comment integration tests ---

func TestInsertComment_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Review this section</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertComment(CommentOptions{
		Text:   "Needs revision.",
		Author: "Reviewer",
		Anchor: "Review this",
	})
	if err != nil {
		t.Fatalf("InsertComment: %v", err)
	}

	// Verify comments.xml created
	cPath := filepath.Join(u.TempDir(), "word", "comments.xml")
	raw, err := os.ReadFile(cPath)
	if err != nil {
		t.Fatalf("read comments.xml: %v", err)
	}
	contents := string(raw)
	if !strings.Contains(contents, "Needs revision.") {
		t.Error("comment text not found")
	}
	if !strings.Contains(contents, "Reviewer") {
		t.Error("comment author not found")
	}

	// Verify comment markers in document
	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "commentRangeStart") {
		t.Error("commentRangeStart not found")
	}
	if !strings.Contains(docXML, "commentRangeEnd") {
		t.Error("commentRangeEnd not found")
	}
	if !strings.Contains(docXML, "commentReference") {
		t.Error("commentReference not found")
	}
}

func TestGetComments_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>text</w:t></w:r></w:p>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	// Insert two comments
	err := u.InsertComment(CommentOptions{Text: "First comment", Author: "Alice", Anchor: "text"})
	if err != nil {
		t.Fatalf("InsertComment 1: %v", err)
	}

	comments, err := u.GetComments()
	if err != nil {
		t.Fatalf("GetComments: %v", err)
	}
	if len(comments) != 1 {
		t.Fatalf("expected 1 comment, got %d", len(comments))
	}
	if comments[0].Author != "Alice" {
		t.Errorf("expected author 'Alice', got %q", comments[0].Author)
	}
}

func TestGetComments_NoComments(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))

	comments, err := u.GetComments()
	if err != nil {
		t.Fatalf("GetComments: %v", err)
	}
	if comments != nil {
		t.Errorf("expected nil for no comments, got %v", comments)
	}
}

// --- Track changes integration tests ---

func TestInsertTrackedText_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Existing content</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTrackedText(TrackedInsertOptions{
		Text:     "New tracked paragraph",
		Author:   "Editor",
		Position: PositionEnd,
		Bold:     true,
	})
	if err != nil {
		t.Fatalf("InsertTrackedText: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "<w:ins") {
		t.Error("tracked insertion not found")
	}
	if !strings.Contains(docXML, "New tracked paragraph") {
		t.Error("tracked text not found")
	}
	if !strings.Contains(docXML, "Editor") {
		t.Error("author not found")
	}
}

func TestInsertTrackedText_EmptyText_Integration(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))
	err := u.InsertTrackedText(TrackedInsertOptions{Text: ""})
	if err == nil {
		t.Error("expected error for empty text")
	}
}

func TestDeleteTrackedText_Integration(t *testing.T) {
	body := `<w:p><w:r><w:t>Keep this paragraph</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Delete this tracked</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.DeleteTrackedText(TrackedDeleteOptions{
		Anchor: "Delete this tracked",
		Author: "Reviewer",
	})
	if err != nil {
		t.Fatalf("DeleteTrackedText: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "<w:del") {
		t.Error("tracked deletion not found")
	}
	if !strings.Contains(docXML, "w:delText") {
		t.Error("deleted text not found")
	}
	if !strings.Contains(docXML, "Reviewer") {
		t.Error("deletion author not found")
	}
}

func TestDeleteTrackedText_EmptyAnchor_Integration(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))
	err := u.DeleteTrackedText(TrackedDeleteOptions{Anchor: ""})
	if err == nil {
		t.Error("expected error for empty anchor")
	}
}

// --- Merge integration tests ---

func TestMergeTableCellsHorizontal_Integration(t *testing.T) {
	body := `<w:tbl>` +
		`<w:tr>` +
		`<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc>` +
		`</w:tr>` +
		`<w:tr>` +
		`<w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc>` +
		`</w:tr>` +
		`</w:tbl>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.MergeTableCellsHorizontal(1, 1, 1, 3)
	if err != nil {
		t.Fatalf("MergeTableCellsHorizontal: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, `w:gridSpan`) {
		t.Error("gridSpan not found after horizontal merge")
	}
	if !strings.Contains(docXML, `w:val="3"`) {
		t.Error("expected gridSpan of 3")
	}
}

func TestMergeTableCellsVertical_Integration(t *testing.T) {
	body := `<w:tbl>` +
		`<w:tr>` +
		`<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>` +
		`</w:tr>` +
		`<w:tr>` +
		`<w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>` +
		`</w:tr>` +
		`<w:tr>` +
		`<w:tc><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc>` +
		`</w:tr>` +
		`</w:tbl>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.MergeTableCellsVertical(1, 1, 3, 1)
	if err != nil {
		t.Fatalf("MergeTableCellsVertical: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, `w:vMerge`) {
		t.Error("vMerge not found after vertical merge")
	}
	if !strings.Contains(docXML, `w:val="restart"`) {
		t.Error("vMerge restart not found")
	}
}

func TestMergeTableCells_TableNotFound(t *testing.T) {
	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, `<w:p/>`))
	err := u.MergeTableCellsHorizontal(1, 1, 1, 2)
	if err == nil {
		t.Error("expected error for non-existent table")
	}
}

// --- io.Reader integration tests ---

func TestNewFromReader_Integration(t *testing.T) {
	fixture := buildIntegrationFixture(t,
		`<w:p><w:r><w:t>Reader test</w:t></w:r></w:p>`+
			`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`)

	reader := bytes.NewReader(fixture)
	u, err := NewFromReader(reader)
	if err != nil {
		t.Fatalf("NewFromReader: %v", err)
	}
	defer u.Cleanup()

	// Verify we can read the document
	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "Reader test") {
		t.Error("document content not readable after NewFromReader")
	}

	// Verify we can make modifications
	err = u.InsertTOC(TOCOptions{
		Title:    "TOC",
		Position: PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertTOC after NewFromReader: %v", err)
	}

	// Verify we can save to writer
	var out bytes.Buffer
	err = u.SaveToWriter(&out)
	if err != nil {
		t.Fatalf("SaveToWriter: %v", err)
	}
	if out.Len() == 0 {
		t.Error("SaveToWriter produced empty output")
	}
}

func TestSaveToWriter_Integration(t *testing.T) {
	fixture := buildIntegrationFixture(t,
		`<w:p><w:r><w:t>Save test</w:t></w:r></w:p>`)

	u := newUpdaterFromFixture(t, fixture)

	var out bytes.Buffer
	err := u.SaveToWriter(&out)
	if err != nil {
		t.Fatalf("SaveToWriter: %v", err)
	}

	// Verify the output is a valid zip
	zipReader, err := zip.NewReader(bytes.NewReader(out.Bytes()), int64(out.Len()))
	if err != nil {
		t.Fatalf("output is not a valid zip: %v", err)
	}

	// Verify it contains document.xml
	found := false
	for _, f := range zipReader.File {
		if f.Name == "word/document.xml" {
			found = true
			break
		}
	}
	if !found {
		t.Error("output zip missing word/document.xml")
	}
}

// --- Additional coverage for edge cases ---

func TestInsertTOC_BeforeAnchor(t *testing.T) {
	body := `<w:p><w:r><w:t>Introduction</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Chapter 1</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTOC(TOCOptions{
		Position: PositionBeforeText,
		Anchor:   "Chapter 1",
	})
	if err != nil {
		t.Fatalf("InsertTOC before: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "fldCharType") {
		t.Error("TOC field not found")
	}
}

func TestInsertMultipleFootnotes(t *testing.T) {
	body := `<w:p><w:r><w:t>First reference</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Second reference</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertFootnote(FootnoteOptions{Text: "Footnote 1", Anchor: "First reference"})
	if err != nil {
		t.Fatalf("InsertFootnote 1: %v", err)
	}

	err = u.InsertFootnote(FootnoteOptions{Text: "Footnote 2", Anchor: "Second reference"})
	if err != nil {
		t.Fatalf("InsertFootnote 2: %v", err)
	}

	fnPath := filepath.Join(u.TempDir(), "word", "footnotes.xml")
	raw, _ := os.ReadFile(fnPath)
	contents := string(raw)
	if !strings.Contains(contents, "Footnote 1") || !strings.Contains(contents, "Footnote 2") {
		t.Error("not all footnotes found")
	}
}

func TestInsertTrackedText_WithFormatting(t *testing.T) {
	body := `<w:p><w:r><w:t>Existing</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTrackedText(TrackedInsertOptions{
		Text:      "Formatted text",
		Author:    "Author",
		Position:  PositionEnd,
		Bold:      true,
		Italic:    true,
		Underline: true,
		Style:     "Heading1",
	})
	if err != nil {
		t.Fatalf("InsertTrackedText: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "<w:b/>") {
		t.Error("bold formatting not found")
	}
	if !strings.Contains(docXML, "<w:i/>") {
		t.Error("italic formatting not found")
	}
	if !strings.Contains(docXML, `<w:u w:val="single"/>`) {
		t.Error("underline formatting not found")
	}
	if !strings.Contains(docXML, `Heading1`) {
		t.Error("style not found")
	}
}

func TestInsertTrackedText_AtBeginning(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.InsertTrackedText(TrackedInsertOptions{
		Text:     "First",
		Author:   "Author",
		Position: PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertTrackedText at beginning: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, "First") {
		t.Error("tracked text not found")
	}
}

func TestSetPageNumber_ReplaceExisting(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgNumType w:start="1" w:fmt="decimal"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	err := u.SetPageNumber(PageNumberOptions{
		Start:  10,
		Format: PageNumLowerRoman,
	})
	if err != nil {
		t.Fatalf("SetPageNumber replace: %v", err)
	}

	docXML := readDocXML(t, u)
	if !strings.Contains(docXML, `w:start="10"`) {
		t.Error("page number start not updated")
	}
	if !strings.Contains(docXML, `w:fmt="lowerRoman"`) {
		t.Error("page number format not updated")
	}
	if strings.Contains(docXML, `w:fmt="decimal"`) {
		t.Error("old page number format still present")
	}
}

func TestDefaultTOCOptions(t *testing.T) {
	opts := DefaultTOCOptions()
	if opts.Title != "Table of Contents" {
		t.Errorf("expected default title, got %q", opts.Title)
	}
	if opts.OutlineLevels != "1-3" {
		t.Errorf("expected default levels 1-3, got %q", opts.OutlineLevels)
	}
	if opts.Position != PositionBeginning {
		t.Errorf("expected default position beginning")
	}
}

func TestDefaultWatermarkOptions(t *testing.T) {
	opts := DefaultWatermarkOptions()
	if opts.Text != "DRAFT" {
		t.Errorf("expected DRAFT, got %q", opts.Text)
	}
	if opts.Opacity != 0.5 {
		t.Errorf("expected opacity 0.5, got %f", opts.Opacity)
	}
	if !opts.Diagonal {
		t.Error("expected diagonal true")
	}
}

func TestInsertComment_DefaultAuthor(t *testing.T) {
	body := `<w:p><w:r><w:t>some text</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	u := newUpdaterFromFixture(t, buildIntegrationFixture(t, body))

	// No author specified - should default to "Author"
	err := u.InsertComment(CommentOptions{
		Text:   "Default author test",
		Anchor: "some text",
	})
	if err != nil {
		t.Fatalf("InsertComment: %v", err)
	}

	cPath := filepath.Join(u.TempDir(), "word", "comments.xml")
	raw, _ := os.ReadFile(cPath)
	if !strings.Contains(string(raw), `w:author="Author"`) {
		t.Error("default author not applied")
	}
}

// Test that saving after modifications works end-to-end
func TestSave_AfterModifications(t *testing.T) {
	body := `<w:p><w:r><w:t>Content</w:t></w:r></w:p>` +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

	path := filepath.Join(t.TempDir(), "input.docx")
	if err := os.WriteFile(path, buildIntegrationFixture(t, body), 0o644); err != nil {
		t.Fatalf("write: %v", err)
	}

	u, err := New(path)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	// Make multiple modifications
	_ = u.InsertTOC(TOCOptions{Title: "TOC", Position: PositionBeginning})
	_ = u.SetPageNumber(PageNumberOptions{Start: 1, Format: PageNumDecimal})
	_ = u.AddStyle(StyleDefinition{ID: "TestStyle", Bold: true})

	// Save
	outPath := filepath.Join(t.TempDir(), "output.docx")
	if err := u.Save(outPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Reopen and verify
	u2, err := New(outPath)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer u2.Cleanup()

	docXML := readDocXML(t, u2)
	if !strings.Contains(docXML, "TOC") {
		t.Error("TOC not persisted")
	}
}

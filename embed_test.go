package godocx_test

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

// buildEmbedFixtureDocx builds a minimal DOCX fixture for embedded-object tests.
func buildEmbedFixtureDocx(t *testing.T) []byte {
	t.Helper()
	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)
	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}
	return docx.Bytes()
}

// minimalXLSX returns the bytes of a minimal valid XLSX (ZIP) file.
func minimalXLSX(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	addZipEntry(t, w, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`)
	addZipEntry(t, w, "_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
	addZipEntry(t, w, "xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>`)
	if err := w.Close(); err != nil {
		t.Fatalf("close xlsx zip: %v", err)
	}
	return buf.Bytes()
}

func TestInsertEmbeddedObjectDefaultIcon(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	xlsxPath := filepath.Join(tempDir, "data.xlsx")
	if err := os.WriteFile(xlsxPath, minimalXLSX(t), 0o644); err != nil {
		t.Fatalf("write xlsx: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
		FilePath: xlsxPath,
		Position: godocx.PositionEnd,
	}); err != nil {
		t.Fatalf("InsertEmbeddedObject: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	entries := listZipEntries(t, outputPath)
	requireEntry(t, entries, "word/embeddings/embedding1.xlsx")
	requireEntry(t, entries, "word/media/image1.png")

	doc := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(doc, "<w:object") {
		t.Error("document.xml missing <w:object element")
	}
	if !strings.Contains(doc, "o:OLEObject") {
		t.Error("document.xml missing o:OLEObject element")
	}
	if !strings.Contains(doc, "Excel.Sheet.12") {
		t.Error("document.xml missing Excel.Sheet.12 ProgID")
	}

	ct := readZipEntry(t, outputPath, "[Content_Types].xml")
	if !strings.Contains(ct, `Extension="xlsx"`) {
		t.Error("[Content_Types].xml missing xlsx extension")
	}

	rels := readZipEntry(t, outputPath, "word/_rels/document.xml.rels")
	if !strings.Contains(rels, "relationships/package") {
		t.Error("document.xml.rels missing package relationship")
	}
	if !strings.Contains(rels, "relationships/image") {
		t.Error("document.xml.rels missing image relationship")
	}
}

func TestInsertEmbeddedObjectFileBytes(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
		FileBytes: minimalXLSX(t),
		FileName:  "report.xlsx",
		Position:  godocx.PositionEnd,
	}); err != nil {
		t.Fatalf("InsertEmbeddedObject: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	entries := listZipEntries(t, outputPath)
	requireEntry(t, entries, "word/embeddings/embedding1.xlsx")
	requireEntry(t, entries, "word/media/image1.png")
}

func TestInsertEmbeddedObjectCustomIcon(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	iconPath := filepath.Join(tempDir, "custom_icon.png")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}
	createTestImage(t, iconPath, 48, 48)

	customIconData, err := os.ReadFile(iconPath)
	if err != nil {
		t.Fatalf("read custom icon: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
		FileBytes: minimalXLSX(t),
		IconPath:  iconPath,
		Position:  godocx.PositionEnd,
	}); err != nil {
		t.Fatalf("InsertEmbeddedObject: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	storedIcon := readZipEntryBytes(t, outputPath, "word/media/image1.png")
	if !bytes.Equal(storedIcon, customIconData) {
		t.Error("stored icon does not match the custom icon provided")
	}
}

func TestInsertEmbeddedObjectMissingIconFallsBack(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	// Provide a non-existent icon path — should fall back silently
	if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
		FileBytes: minimalXLSX(t),
		IconPath:  filepath.Join(tempDir, "does_not_exist.png"),
		Position:  godocx.PositionEnd,
	}); err != nil {
		t.Fatalf("InsertEmbeddedObject should not fail on bad IconPath, got: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Built-in icon must be present
	entries := listZipEntries(t, outputPath)
	requireEntry(t, entries, "word/media/image1.png")
}

func TestInsertEmbeddedObjectMultiple(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	for i := 0; i < 2; i++ {
		if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
			FileBytes: minimalXLSX(t),
			Position:  godocx.PositionEnd,
		}); err != nil {
			t.Fatalf("InsertEmbeddedObject #%d: %v", i+1, err)
		}
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	entries := listZipEntries(t, outputPath)
	requireEntry(t, entries, "word/embeddings/embedding1.xlsx")
	requireEntry(t, entries, "word/embeddings/embedding2.xlsx")
	requireEntry(t, entries, "word/media/image1.png")
	requireEntry(t, entries, "word/media/image2.png")

	doc := readZipEntry(t, outputPath, "word/document.xml")
	count := strings.Count(doc, "<w:object")
	if count != 2 {
		t.Errorf("expected 2 <w:object elements, got %d", count)
	}
}

func TestInsertEmbeddedObjectPositions(t *testing.T) {
	tests := []struct {
		name     string
		position godocx.InsertPosition
		anchor   string
	}{
		{"End", godocx.PositionEnd, ""},
		{"Beginning", godocx.PositionBeginning, ""},
		{"AfterText", godocx.PositionAfterText, "Hello World"},
		{"BeforeText", godocx.PositionBeforeText, "Hello World"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tempDir := t.TempDir()
			inputPath := filepath.Join(tempDir, "input.docx")
			outputPath := filepath.Join(tempDir, "output.docx")

			if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
				t.Fatalf("write input: %v", err)
			}

			u, err := godocx.New(inputPath)
			if err != nil {
				t.Fatalf("New: %v", err)
			}
			defer u.Cleanup()

			if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
				FileBytes: minimalXLSX(t),
				Position:  tt.position,
				Anchor:    tt.anchor,
			}); err != nil {
				t.Fatalf("InsertEmbeddedObject: %v", err)
			}

			if err := u.Save(outputPath); err != nil {
				t.Fatalf("Save: %v", err)
			}

			doc := readZipEntry(t, outputPath, "word/document.xml")
			if !strings.Contains(doc, "<w:object") {
				t.Error("document.xml missing <w:object")
			}
		})
	}
}

func TestInsertEmbeddedObjectValidation(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	// Neither FilePath nor FileBytes set → error
	err = u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{Position: godocx.PositionEnd})
	if err == nil {
		t.Error("expected error when neither FilePath nor FileBytes is set")
	}
}

func TestInsertEmbeddedObjectAnchorRequired(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	// PositionAfterText without Anchor → error
	err = u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
		FileBytes: minimalXLSX(t),
		Position:  godocx.PositionAfterText,
	})
	if err == nil {
		t.Error("expected error when Anchor is empty for PositionAfterText")
	}
}

func TestInsertEmbeddedObjectDrawAspectIcon(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildEmbedFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	if err := u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
		FileBytes: minimalXLSX(t),
		Position:  godocx.PositionEnd,
	}); err != nil {
		t.Fatalf("InsertEmbeddedObject: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	doc := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(doc, `DrawAspect="Icon"`) {
		t.Error("document.xml missing DrawAspect=\"Icon\" attribute - object will not display as icon")
	}
	// Object must be embedded (not linked): Type="Embed" and no Type="Link".
	if !strings.Contains(doc, `Type="Embed"`) {
		t.Error("document.xml missing Type=\"Embed\" - object is not embedded")
	}
	if strings.Contains(doc, `Type="Link"`) {
		t.Error("document.xml contains Type=\"Link\" - object must not be a linked file")
	}
	// No hyperlink wrapper around the OLE object.
	if strings.Contains(doc, "<w:hyperlink") {
		t.Error("document.xml contains <w:hyperlink> around embedded object - icon must be inserted without a link")
	}
}

// requireEntry asserts that the named entry exists in the entries slice.
func requireEntry(t *testing.T, entries []string, name string) {
	t.Helper()
	for _, e := range entries {
		if e == name {
			return
		}
	}
	t.Errorf("ZIP entry %q not found; entries: %v", name, entries)
}

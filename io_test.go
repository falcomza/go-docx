package godocx

import (
	"archive/zip"
	"bytes"
	"os"
	"strings"
	"testing"
)

func TestNewFromReader_NilReader(t *testing.T) {
	_, err := NewFromReader(nil)
	if err == nil {
		t.Error("expected error for nil reader")
	}
}

func TestSaveToWriter_NilUpdater(t *testing.T) {
	var u *Updater
	var buf bytes.Buffer
	err := u.SaveToWriter(&buf)
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestSaveToWriter_NilWriter(t *testing.T) {
	// We need a valid Updater for this test; create a minimal temp dir
	tmpDir, err := os.MkdirTemp("", "docx-test-*")
	if err != nil {
		t.Fatalf("create temp dir: %v", err)
	}
	defer os.RemoveAll(tmpDir)

	u := &Updater{tempDir: tmpDir}
	err = u.SaveToWriter(nil)
	if err == nil {
		t.Error("expected error for nil writer")
	}
}

// buildFixtureDotx returns a minimal in-memory .dotx ZIP with the template content type.
func buildFixtureDotx(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZipEntry(t, zw, "[Content_Types].xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`+
			`<Override PartName="/word/document.xml" ContentType="`+DotxMainContentType+`"/>`+
			`</Types>`)
	addZipEntry(t, zw, "word/document.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`+
			`<w:body><w:p><w:r><w:t>Hello from template</w:t></w:r></w:p></w:body>`+
			`</w:document>`)
	addZipEntry(t, zw, "word/_rels/document.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
	if err := zw.Close(); err != nil {
		t.Fatalf("close dotx zip: %v", err)
	}
	return buf.Bytes()
}

func TestNew_DotxPromotesContentType(t *testing.T) {
	tmp := t.TempDir()
	dotxPath := tmp + "/template.dotx"
	if err := os.WriteFile(dotxPath, buildFixtureDotx(t), 0o644); err != nil {
		t.Fatalf("write dotx: %v", err)
	}

	u, err := New(dotxPath)
	if err != nil {
		t.Fatalf("New(.dotx) failed: %v", err)
	}
	defer u.Cleanup()

	ct := readZipEntry(t, dotxPath, "[Content_Types].xml")
	// The source still has the template type — only the in-memory copy is promoted.
	if !strings.Contains(ct, DotxMainContentType) {
		t.Fatal("source fixture unexpectedly missing template content type")
	}

	// Verify the extracted (working) copy has been promoted to document type.
	extracted, err := os.ReadFile(u.TempDir() + "/[Content_Types].xml")
	if err != nil {
		t.Fatalf("read extracted [Content_Types].xml: %v", err)
	}
	extractedStr := string(extracted)
	if strings.Contains(extractedStr, DotxMainContentType) {
		t.Error("template content type was not replaced")
	}
	if !strings.Contains(extractedStr, DocxMainContentType) {
		t.Error("document content type not found after promotion")
	}
}

func TestNew_DotxSavesAsDocx(t *testing.T) {
	tmp := t.TempDir()
	dotxPath := tmp + "/template.dotx"
	outputPath := tmp + "/output.docx"

	if err := os.WriteFile(dotxPath, buildFixtureDotx(t), 0o644); err != nil {
		t.Fatalf("write dotx: %v", err)
	}

	u, err := New(dotxPath)
	if err != nil {
		t.Fatalf("New(.dotx) failed: %v", err)
	}
	defer u.Cleanup()

	if err := u.InsertParagraph(ParagraphOptions{Text: "Added paragraph", Position: PositionEnd}); err != nil {
		t.Fatalf("InsertParagraph: %v", err)
	}
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	ct := readZipEntry(t, outputPath, "[Content_Types].xml")
	if strings.Contains(ct, DotxMainContentType) {
		t.Error("output docx still contains template content type")
	}
	if !strings.Contains(ct, DocxMainContentType) {
		t.Error("output docx missing document content type")
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Hello from template") {
		t.Error("original template text not found in output")
	}
	if !strings.Contains(docXML, "Added paragraph") {
		t.Error("inserted paragraph not found in output")
	}
}

func TestNew_DotxViaNewFromBytes(t *testing.T) {
	u, err := NewFromBytes(buildFixtureDotx(t))
	if err != nil {
		t.Fatalf("NewFromBytes(.dotx) failed: %v", err)
	}
	defer u.Cleanup()

	extracted, err := os.ReadFile(u.TempDir() + "/[Content_Types].xml")
	if err != nil {
		t.Fatalf("read [Content_Types].xml: %v", err)
	}
	if strings.Contains(string(extracted), DotxMainContentType) {
		t.Error("template content type was not replaced via NewFromBytes")
	}
}

// buildFixtureNS0Docx returns a minimal in-memory DOCX where the
// WordprocessingML namespace is bound to the "ns0" prefix instead of the
// canonical "w" prefix.  This mimics the output of python-docx / lxml.
func buildFixtureNS0Docx(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZipEntry(t, zw, "[Content_Types].xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`+
			`<Override PartName="/word/document.xml" ContentType="`+DocxMainContentType+`"/>`+
			`</Types>`)
	addZipEntry(t, zw, "_rels/.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`+
			`</Relationships>`)
	// document.xml uses "ns0:" for WML elements — the python-docx / lxml style
	addZipEntry(t, zw, "word/document.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<ns0:document xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`+
			`<ns0:body>`+
			`<ns0:p><ns0:r><ns0:t>Template content</ns0:t></ns0:r></ns0:p>`+
			`<ns0:sectPr><ns0:pgSz ns0:w="12240" ns0:h="15840"/></ns0:sectPr>`+
			`</ns0:body>`+
			`</ns0:document>`)
	addZipEntry(t, zw, "word/_rels/document.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
	if err := zw.Close(); err != nil {
		t.Fatalf("close ns0 docx zip: %v", err)
	}
	return buf.Bytes()
}

// TestNew_NS0NamespacePrefixNormalised verifies that a DOCX produced by
// python-docx/lxml — which binds the WML namespace to "ns0:" instead of
// the canonical "w:" — can have content appended with AddHeading,
// InsertParagraph, and InsertTable, and that the appended content is
// present in the saved output.
//
// Regression test for: appended content missing when template uses a
// non-standard WML namespace prefix (the "ns0 bug").
func TestNew_NS0NamespacePrefixNormalised(t *testing.T) {
	tmp := t.TempDir()
	docxPath := tmp + "/ns0_template.docx"
	outputPath := tmp + "/ns0_output.docx"

	if err := os.WriteFile(docxPath, buildFixtureNS0Docx(t), 0o644); err != nil {
		t.Fatalf("write ns0 docx: %v", err)
	}

	u, err := New(docxPath)
	if err != nil {
		t.Fatalf("New(ns0 docx) failed: %v", err)
	}
	defer u.Cleanup()

	// Verify document.xml was normalised to canonical w: prefix.
	extracted, err := os.ReadFile(u.TempDir() + "/word/document.xml")
	if err != nil {
		t.Fatalf("read extracted document.xml: %v", err)
	}
	extractedStr := string(extracted)
	if strings.Contains(extractedStr, "ns0:") {
		t.Error("document.xml still contains non-canonical ns0: prefix after normalisation")
	}
	if !strings.Contains(extractedStr, "xmlns:w=") {
		t.Error("document.xml missing canonical xmlns:w declaration after normalisation")
	}

	// Append content using all three common methods.
	if err := u.AddHeading(1, "AppendedHeading", PositionEnd); err != nil {
		t.Fatalf("AddHeading: %v", err)
	}
	if err := u.InsertParagraph(ParagraphOptions{Text: "AppendedParagraph", Position: PositionEnd}); err != nil {
		t.Fatalf("InsertParagraph: %v", err)
	}
	if err := u.InsertTable(TableOptions{
		Position: PositionEnd,
		Columns:  []ColumnDefinition{{Title: "ColA"}, {Title: "ColB"}},
		Rows:     [][]string{{"Cell1", "Cell2"}},
	}); err != nil {
		t.Fatalf("InsertTable: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")

	for _, want := range []string{"Template content", "AppendedHeading", "AppendedParagraph", "Cell1", "Cell2"} {
		if !strings.Contains(docXML, want) {
			t.Errorf("saved document.xml missing expected text %q", want)
		}
	}
}

func TestWriteZipFromDir(t *testing.T) {
	// Create a minimal directory structure
	tmpDir, err := os.MkdirTemp("", "zip-test-*")
	if err != nil {
		t.Fatalf("create temp dir: %v", err)
	}
	defer os.RemoveAll(tmpDir)

	// Write a test file
	if err := os.WriteFile(tmpDir+"/test.txt", []byte("hello"), 0o644); err != nil {
		t.Fatalf("write test file: %v", err)
	}

	var buf bytes.Buffer
	if err := writeZipFromDir(tmpDir, &buf); err != nil {
		t.Fatalf("writeZipFromDir: %v", err)
	}

	if buf.Len() == 0 {
		t.Error("expected non-empty zip output")
	}
}

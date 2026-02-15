package docxupdater_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

func TestInsertParagraphAtEnd(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert a paragraph at the end
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This is a test paragraph",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify the paragraph was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "This is a test paragraph") {
		t.Error("Paragraph text not found in document.xml")
	}
}

func TestInsertParagraphAtBeginning(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Beginning paragraph",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Beginning paragraph") {
		t.Error("Paragraph text not found in document.xml")
	}
	if !strings.Contains(docXML, "Heading1") {
		t.Error("Heading1 style not found in document.xml")
	}
}

func TestInsertParagraphWithFormatting(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:      "Bold and italic text",
		Style:     docxupdater.StyleNormal,
		Position:  docxupdater.PositionEnd,
		Bold:      true,
		Italic:    true,
		Underline: true,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:b/>") {
		t.Error("Bold formatting not found")
	}
	if !strings.Contains(docXML, "<w:i/>") {
		t.Error("Italic formatting not found")
	}
	if !strings.Contains(docXML, "<w:u") {
		t.Error("Underline formatting not found")
	}
}

func TestAddHeading(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	if err := u.AddHeading(1, "Main Title", docxupdater.PositionEnd); err != nil {
		t.Fatalf("AddHeading failed: %v", err)
	}

	if err := u.AddHeading(2, "Subtitle", docxupdater.PositionEnd); err != nil {
		t.Fatalf("AddHeading failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Main Title") {
		t.Error("Heading 1 text not found")
	}
	if !strings.Contains(docXML, "Subtitle") {
		t.Error("Heading 2 text not found")
	}
}

func TestInsertMultipleParagraphs(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	paragraphs := []docxupdater.ParagraphOptions{
		{
			Text:     "First paragraph",
			Style:    docxupdater.StyleHeading1,
			Position: docxupdater.PositionEnd,
		},
		{
			Text:     "Second paragraph with details",
			Style:    docxupdater.StyleNormal,
			Position: docxupdater.PositionEnd,
		},
		{
			Text:     "Third paragraph conclusion",
			Style:    docxupdater.StyleNormal,
			Position: docxupdater.PositionEnd,
			Bold:     true,
		},
	}

	if err := u.InsertParagraphs(paragraphs); err != nil {
		t.Fatalf("InsertParagraphs failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	for _, para := range paragraphs {
		if !strings.Contains(docXML, para.Text) {
			t.Errorf("Paragraph text %q not found in document", para.Text)
		}
	}
}

func TestInsertParagraphEmptyText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "",
		Position: docxupdater.PositionEnd,
	})
	if err == nil {
		t.Error("Expected error for empty text, got nil")
	}
}

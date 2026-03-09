package godocx

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestListParagraphStyleApplied(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add bullet list items
	err = u.AddBulletItem("First bullet point", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletItem failed: %v", err)
	}

	err = u.AddBulletItem("Second bullet point", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletItem failed: %v", err)
	}

	// Add numbered list items
	err = u.AddNumberedItem("First numbered item", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddNumberedItem failed: %v", err)
	}

	err = u.AddNumberedItem("Second numbered item", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddNumberedItem failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify styles.xml was created/updated with ListParagraph style
	stylesXML := readZipEntry(t, outputPath, "word/styles.xml")
	if stylesXML == "" {
		t.Fatal("styles.xml not found")
	}

	// Verify ListParagraph style is defined
	if !strings.Contains(stylesXML, `w:styleId="ListParagraph"`) {
		t.Error("ListParagraph style not defined in styles.xml")
	}

	if !strings.Contains(stylesXML, `<w:name w:val="List Paragraph"/>`) {
		t.Error("ListParagraph style name not found in styles.xml")
	}

	// Verify document.xml applies ListParagraph style to list items
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if docXML == "" {
		t.Fatal("document.xml not found")
	}

	// Check that list paragraphs have both the style and numbering properties
	// There should be paragraphs with both <w:pStyle w:val="ListParagraph"/> and <w:numPr>
	if !strings.Contains(docXML, `<w:pStyle w:val="ListParagraph"/>`) {
		t.Error("ListParagraph style not applied to list items in document.xml")
	}

	// Count how many list items have the style applied
	styleCount := strings.Count(docXML, `<w:pStyle w:val="ListParagraph"/>`)
	numPrCount := strings.Count(docXML, `<w:numPr>`)

	// We added 4 list items (2 bullet + 2 numbered), so there should be 4 list style applications
	if styleCount != 4 {
		t.Errorf("Expected 4 ListParagraph style applications, got %d", styleCount)
	}

	// And 4 numbering property blocks
	if numPrCount != 4 {
		t.Errorf("Expected 4 numPr blocks, got %d", numPrCount)
	}
}

func TestListParagraphStyleInMinimalDocument(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "minimal_list.docx")

	// Create input document from fixture
	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	// Create a new document from fixture
	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add a single bullet item
	err = u.AddBulletItem("Test bullet", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletItem failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify styles.xml contains ListParagraph style
	stylesXML := readZipEntry(t, outputPath, "word/styles.xml")
	if !strings.Contains(stylesXML, `w:styleId="ListParagraph"`) {
		t.Error("ListParagraph style not found in minimal document's styles.xml")
	}

	// Verify document.xml applies the style
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:pStyle w:val="ListParagraph"/>`) {
		t.Error("ListParagraph style not applied in minimal document")
	}
}

func TestInsertParagraphWithListAppliesStyle(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Use InsertParagraph with ListType
	err = u.InsertParagraph(ParagraphOptions{
		Text:      "Bullet item via InsertParagraph",
		ListType:  ListTypeBullet,
		ListLevel: 0,
		Position:  PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	err = u.InsertParagraph(ParagraphOptions{
		Text:      "Numbered item via InsertParagraph",
		ListType:  ListTypeNumbered,
		ListLevel: 0,
		Position:  PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify ListParagraph style is applied
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	styleCount := strings.Count(docXML, `<w:pStyle w:val="ListParagraph"/>`)
	if styleCount != 2 {
		t.Errorf("Expected 2 ListParagraph style applications via InsertParagraph, got %d", styleCount)
	}
}

func TestExplicitStyleOverridesListParagraph(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Use InsertParagraph with explicit style that is NOT ListParagraph
	err = u.InsertParagraph(ParagraphOptions{
		Text:      "Custom styled list item",
		Style:     "Heading1",
		ListType:  ListTypeBullet,
		ListLevel: 0,
		Position:  PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify that explicit style is used instead of ListParagraph
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:pStyle w:val="Heading1"/>`) {
		t.Error("Explicit style Heading1 not applied")
	}

	// Should NOT have ListParagraph since we explicitly set a different style
	if strings.Contains(docXML, `<w:pStyle w:val="ListParagraph"/>`) {
		t.Error("ListParagraph should not be applied when explicit style is set")
	}
}

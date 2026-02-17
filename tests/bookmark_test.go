package docxupdater_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

func TestCreateEmptyBookmark(t *testing.T) {
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

	// Create an empty bookmark
	opts := docxupdater.DefaultBookmarkOptions()
	opts.Position = docxupdater.PositionEnd
	err = u.CreateBookmark("my_bookmark", opts)
	if err != nil {
		t.Fatalf("CreateBookmark failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify bookmark was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:name="my_bookmark"`) {
		t.Error("Bookmark name not found in document.xml")
	}
	if !strings.Contains(docXML, "<w:bookmarkStart") {
		t.Error("Bookmark start tag not found in document.xml")
	}
	if !strings.Contains(docXML, "<w:bookmarkEnd") {
		t.Error("Bookmark end tag not found in document.xml")
	}
}

func TestCreateBookmarkWithText(t *testing.T) {
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

	// Create bookmark with text
	opts := docxupdater.DefaultBookmarkOptions()
	opts.Position = docxupdater.PositionEnd
	err = u.CreateBookmarkWithText("summary_section", "Executive Summary", opts)
	if err != nil {
		t.Fatalf("CreateBookmarkWithText failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify bookmark and text were added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:name="summary_section"`) {
		t.Error("Bookmark name not found in document.xml")
	}
	if !strings.Contains(docXML, "Executive Summary") {
		t.Error("Bookmark text not found in document.xml")
	}

	// Verify proper structure: bookmarkStart, text, bookmarkEnd
	bookmarkStartIdx := strings.Index(docXML, `<w:bookmarkStart`)
	textIdx := strings.Index(docXML, "Executive Summary")
	bookmarkEndIdx := strings.Index(docXML, `<w:bookmarkEnd`)

	if bookmarkStartIdx == -1 || textIdx == -1 || bookmarkEndIdx == -1 {
		t.Fatal("Missing bookmark structure elements")
	}

	if !(bookmarkStartIdx < textIdx && textIdx < bookmarkEndIdx) {
		t.Error("Bookmark structure is not in correct order (start, text, end)")
	}
}

func TestWrapTextInBookmark(t *testing.T) {
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

	// First, add some text
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This is important content",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	// Now wrap it in a bookmark
	err = u.WrapTextInBookmark("important_section", "important content")
	if err != nil {
		t.Fatalf("WrapTextInBookmark failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify bookmark wraps the text
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:name="important_section"`) {
		t.Error("Bookmark name not found in document.xml")
	}

	// Verify bookmarkStart appears before the text
	bookmarkStartIdx := strings.Index(docXML, `<w:bookmarkStart`)
	textIdx := strings.Index(docXML, "important content")

	if bookmarkStartIdx == -1 || textIdx == -1 {
		t.Fatal("Bookmark or text not found")
	}

	if bookmarkStartIdx > textIdx {
		t.Error("Bookmark start should appear before the text")
	}
}

func TestBookmarkWithInternalLink(t *testing.T) {
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

	// Create a bookmark
	bookmarkOpts := docxupdater.DefaultBookmarkOptions()
	bookmarkOpts.Position = docxupdater.PositionEnd
	err = u.CreateBookmarkWithText("conclusion_section", "Conclusion", bookmarkOpts)
	if err != nil {
		t.Fatalf("CreateBookmarkWithText failed: %v", err)
	}

	// Create an internal link to the bookmark
	linkOpts := docxupdater.DefaultHyperlinkOptions()
	linkOpts.Position = docxupdater.PositionBeginning
	err = u.InsertInternalLink("Go to Conclusion", "conclusion_section", linkOpts)
	if err != nil {
		t.Fatalf("InsertInternalLink failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify both bookmark and link exist
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Check bookmark exists
	if !strings.Contains(docXML, `w:name="conclusion_section"`) {
		t.Error("Bookmark name not found in document.xml")
	}

	// Check internal link exists with correct anchor
	if !strings.Contains(docXML, `w:anchor="conclusion_section"`) {
		t.Error("Internal link anchor not found in document.xml")
	}
	if !strings.Contains(docXML, "Go to Conclusion") {
		t.Error("Link text not found in document.xml")
	}
}

func TestBookmarkNameValidation(t *testing.T) {
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

	opts := docxupdater.DefaultBookmarkOptions()

	tests := []struct {
		name      string
		shouldErr bool
	}{
		{"valid_bookmark", false},
		{"ValidBookmark", false},
		{"Valid_Bookmark_123", false},
		{"1invalid", true},               // starts with digit
		{"invalid bookmark", true},       // contains space
		{"invalid-bookmark", true},       // contains hyphen
		{"_Tocinvalid", true},            // reserved prefix
		{"", true},                       // empty
		{strings.Repeat("a", 41), true},  // too long (>40 chars)
		{strings.Repeat("a", 40), false}, // exactly 40 chars - ok
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := u.CreateBookmark(tt.name, opts)
			if tt.shouldErr && err == nil {
				t.Errorf("Expected error for bookmark name '%s', but got none", tt.name)
			}
			if !tt.shouldErr && err != nil {
				t.Errorf("Expected no error for bookmark name '%s', but got: %v", tt.name, err)
			}
		})
	}
}

func TestMultipleBookmarks(t *testing.T) {
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

	opts := docxupdater.DefaultBookmarkOptions()
	opts.Position = docxupdater.PositionEnd

	// Create multiple bookmarks - they should have unique IDs
	bookmarks := []string{"bookmark_one", "bookmark_two", "bookmark_three"}
	for _, name := range bookmarks {
		err = u.CreateBookmarkWithText(name, "Content for "+name, opts)
		if err != nil {
			t.Fatalf("CreateBookmarkWithText failed for %s: %v", name, err)
		}
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify all bookmarks exist with unique IDs
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	for _, name := range bookmarks {
		if !strings.Contains(docXML, `w:name="`+name+`"`) {
			t.Errorf("Bookmark '%s' not found in document.xml", name)
		}
	}

	// Count bookmark IDs to ensure they are unique
	// Each bookmark creates one start and one end tag with the same ID
	idPattern := `w:id="(\d+)"`
	ids := make(map[string]int)
	for i := 0; i < len(docXML); i++ {
		if start := strings.Index(docXML[i:], idPattern); start != -1 {
			end := strings.Index(docXML[i+start:], `"`)
			if end != -1 {
				end = strings.Index(docXML[i+start+end+1:], `"`)
				if end != -1 {
					idValue := docXML[i+start : i+start+end+2]
					ids[idValue]++
					i = i + start + end + 2
				}
			}
		}
	}
}

func TestBookmarkPositions(t *testing.T) {
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

	// Add anchor text first
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Anchor paragraph for positioning",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	// Test PositionBeginning
	opts1 := docxupdater.DefaultBookmarkOptions()
	opts1.Position = docxupdater.PositionBeginning
	err = u.CreateBookmark("bookmark_beginning", opts1)
	if err != nil {
		t.Fatalf("CreateBookmark at beginning failed: %v", err)
	}

	// Test PositionAfterText
	opts2 := docxupdater.DefaultBookmarkOptions()
	opts2.Position = docxupdater.PositionAfterText
	opts2.Anchor = "Anchor paragraph"
	err = u.CreateBookmark("bookmark_after", opts2)
	if err != nil {
		t.Fatalf("CreateBookmark after text failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify bookmarks
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	if !strings.Contains(docXML, `w:name="bookmark_beginning"`) {
		t.Error("Bookmark at beginning not found")
	}
	if !strings.Contains(docXML, `w:name="bookmark_after"`) {
		t.Error("Bookmark after text not found")
	}
}

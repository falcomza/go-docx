package docxupdater_test

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update"
)

// createTestImage creates a simple test image with the given dimensions
func createTestImage(t *testing.T, path string, width, height int) {
	t.Helper()

	// Create a simple colored image
	img := image.NewRGBA(image.Rect(0, 0, width, height))
	for y := range height {
		for x := range width {
			img.Set(x, y, color.RGBA{uint8(x % 256), uint8(y % 256), 128, 255})
		}
	}

	// Ensure directory exists
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatalf("create image dir: %v", err)
	}

	// Save as PNG
	f, err := os.Create(path)
	if err != nil {
		t.Fatalf("create image file: %v", err)
	}
	defer f.Close()

	if err := png.Encode(f, img); err != nil {
		t.Fatalf("encode PNG: %v", err)
	}
}

func TestInsertImageWithWidthOnly(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	// Create test image (800x600)
	createTestImage(t, imagePath, 800, 600)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert image with only width specified (400px)
	// Expected proportional height: 400 * (600/800) = 300px
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Width:    400,
		AltText:  "Test Image",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify document contains the image
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Test Image") {
		t.Error("Image alt text not found in document.xml")
	}
	if !strings.Contains(docXML, "pic:pic") {
		t.Error("Image drawing not found in document.xml")
	}

	// Verify relationship exists
	relsXML := readZipEntry(t, outputPath, "word/_rels/document.xml.rels")
	if !strings.Contains(relsXML, "media/image1.png") {
		t.Error("Image relationship not found")
	}
}

func TestInsertImageWithHeightOnly(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	// Create test image (1200x800)
	createTestImage(t, imagePath, 1200, 800)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert image with only height specified (400px)
	// Expected proportional width: 400 * (1200/800) = 600px
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Height:   400,
		AltText:  "Test Image Height",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify document contains the image
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Test Image Height") {
		t.Error("Image alt text not found in document.xml")
	}
	if !strings.Contains(docXML, "pic:pic") {
		t.Error("Image drawing not found in document.xml")
	}
}

func TestInsertImageWithBothDimensions(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	// Create test image
	createTestImage(t, imagePath, 800, 600)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert image with both width and height specified
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Width:    500,
		Height:   300,
		AltText:  "Test Image Both",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify document contains the image
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Test Image Both") {
		t.Error("Image alt text not found in document.xml")
	}
}

func TestInsertImageWithNoDimensions(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	// Create test image (640x480)
	createTestImage(t, imagePath, 640, 480)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert image with no dimensions (use actual size)
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		AltText:  "Test Image Actual Size",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify document contains the image
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Test Image Actual Size") {
		t.Error("Image alt text not found in document.xml")
	}
}

func TestInsertImageAtBeginning(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	createTestImage(t, imagePath, 400, 300)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert at beginning
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Width:    300,
		AltText:  "Beginning Image",
		Position: docxupdater.PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Beginning Image") {
		t.Error("Image not found at beginning")
	}
}

func TestInsertMultipleImages(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	image1Path := filepath.Join(tempDir, "test_image1.png")
	image2Path := filepath.Join(tempDir, "test_image2.png")

	createTestImage(t, image1Path, 800, 600)
	createTestImage(t, image2Path, 1024, 768)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert first image
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image1Path,
		Width:    400,
		AltText:  "First Image",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage 1 failed: %v", err)
	}

	// Insert second image
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image2Path,
		Width:    500,
		AltText:  "Second Image",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage 2 failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify both images are present
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "First Image") {
		t.Error("First image not found")
	}
	if !strings.Contains(docXML, "Second Image") {
		t.Error("Second image not found")
	}

	// Verify relationships
	relsXML := readZipEntry(t, outputPath, "word/_rels/document.xml.rels")
	if !strings.Contains(relsXML, "media/image1.png") {
		t.Error("First image relationship not found")
	}
	if !strings.Contains(relsXML, "media/image2.png") {
		t.Error("Second image relationship not found")
	}
}

func TestInsertImageAfterText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	createTestImage(t, imagePath, 600, 400)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add anchor text
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Insert image after this text",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	// Insert image after text
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Width:    400,
		AltText:  "After Text Image",
		Position: docxupdater.PositionAfterText,
		Anchor:   "Insert image after",
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "After Text Image") {
		t.Error("Image not found")
	}

	// Check that image appears after the anchor text
	anchorPos := bytes.Index([]byte(docXML), []byte("Insert image after"))
	imagePos := bytes.Index([]byte(docXML), []byte("After Text Image"))
	if anchorPos == -1 || imagePos == -1 {
		t.Error("Anchor or image text not found")
	} else if imagePos < anchorPos {
		t.Error("Image appears before anchor text, expected after")
	}
}

func TestInsertImageBeforeText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	createTestImage(t, imagePath, 500, 400)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add anchor text
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Insert image before this text",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	// Insert image before text
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Height:   300,
		AltText:  "Before Text Image",
		Position: docxupdater.PositionBeforeText,
		Anchor:   "Insert image before",
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Before Text Image") {
		t.Error("Image not found")
	}

	// Check that image appears before the anchor text
	anchorPos := bytes.Index([]byte(docXML), []byte("Insert image before"))
	imagePos := bytes.Index([]byte(docXML), []byte("Before Text Image"))
	if anchorPos == -1 || imagePos == -1 {
		t.Error("Anchor or image text not found")
	} else if imagePos > anchorPos {
		t.Error("Image appears after anchor text, expected before")
	}
}

func TestInsertImageInvalidPath(t *testing.T) {
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

	// Try to insert non-existent image
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     "nonexistent.png",
		Position: docxupdater.PositionEnd,
	})
	if err == nil {
		t.Error("Expected error for non-existent image, got nil")
	}
}

func TestInsertImageEmptyPath(t *testing.T) {
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

	// Try to insert with empty path
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     "",
		Position: docxupdater.PositionEnd,
	})
	if err == nil {
		t.Error("Expected error for empty path, got nil")
	}
}

func TestContentTypeRegistration(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")
	imagePath := filepath.Join(tempDir, "test_image.png")

	createTestImage(t, imagePath, 400, 300)

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     imagePath,
		Width:    300,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertImage failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify content type is registered
	contentTypesXML := readZipEntry(t, outputPath, "[Content_Types].xml")
	if !strings.Contains(contentTypesXML, "image/png") {
		t.Error("PNG content type not registered")
	}
}

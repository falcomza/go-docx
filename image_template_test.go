package docxupdater_test

import (
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update"
)

// TestInsertImageWithRealTemplate tests image insertion using the actual docx_template.docx
func TestInsertImageWithRealTemplate(t *testing.T) {
	// Setup paths
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_image_test.docx")

	// Create test image
	tempDir := t.TempDir()
	testImagePath := filepath.Join(tempDir, "test_image.png")
	createTestImageForTemplate(t, testImagePath, 800, 600)

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	// Open the template
	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Insert image with proportional width
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Width:    400, // Height will be calculated as 300
		AltText:  "Test Image from Template",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert image: %v", err)
	}

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the output file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Verify document contains the image
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Test Image from Template") {
		t.Error("Image alt text not found in document.xml")
	}
	if !strings.Contains(docXML, "pic:pic") {
		t.Error("Image drawing not found in document.xml")
	}

	// Verify relationship exists
	relsXML := readZipEntry(t, outputPath, "word/_rels/document.xml.rels")
	if !strings.Contains(relsXML, "media/image") {
		t.Error("Image relationship not found")
	}

	t.Logf("Successfully created document with image at: %s", outputPath)
}

// TestInsertMultipleImagesInRealTemplate tests inserting multiple images
func TestInsertMultipleImagesInRealTemplate(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_multiple_images_test.docx")

	// Create test images
	tempDir := t.TempDir()
	image1Path := filepath.Join(tempDir, "image1.png")
	image2Path := filepath.Join(tempDir, "image2.png")
	createTestImageForTemplate(t, image1Path, 640, 480)
	createTestImageForTemplate(t, image2Path, 1024, 768)

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Add a heading
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Image Gallery Test",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert heading: %v", err)
	}

	// Insert first image with width only
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image1Path,
		Width:    400,
		AltText:  "First Test Image",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert first image: %v", err)
	}

	// Add separator text
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Figure 2:",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
		Bold:     true,
	})
	if err != nil {
		t.Fatalf("Failed to insert separator: %v", err)
	}

	// Insert second image with height only
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image2Path,
		Height:   350,
		AltText:  "Second Test Image",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert second image: %v", err)
	}

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the output file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Verify document contains both images
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "First Test Image") {
		t.Error("First image alt text not found")
	}
	if !strings.Contains(docXML, "Second Test Image") {
		t.Error("Second image alt text not found")
	}

	// Verify relationships
	relsXML := readZipEntry(t, outputPath, "word/_rels/document.xml.rels")
	if !strings.Contains(relsXML, "media/image1.png") {
		t.Error("First image relationship not found")
	}
	if !strings.Contains(relsXML, "media/image2.png") {
		t.Error("Second image relationship not found")
	}

	t.Logf("Successfully created document with multiple images at: %s", outputPath)
}

// TestInsertImageWithTextAnchorsInRealTemplate tests position-based insertion
func TestInsertImageWithTextAnchorsInRealTemplate(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_anchored_images_test.docx")

	// Create test image
	tempDir := t.TempDir()
	testImagePath := filepath.Join(tempDir, "diagram.png")
	createTestImageForTemplate(t, testImagePath, 800, 500)

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Add anchor text at the beginning
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Introduction Section",
		Style:    docxupdater.StyleHeading2,
		Position: docxupdater.PositionBeginning,
	})
	if err != nil {
		t.Fatalf("Failed to insert anchor heading: %v", err)
	}

	// Insert image after the introduction heading
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Width:    500,
		AltText:  "Introduction Diagram",
		Position: docxupdater.PositionAfterText,
		Anchor:   "Introduction Section",
	})
	if err != nil {
		t.Fatalf("Failed to insert image after text: %v", err)
	}

	// Add more text
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Figure 1: System Overview",
		Style:    docxupdater.StyleHeading3,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert figure caption: %v", err)
	}

	// Insert image before the figure caption
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Height:   300,
		AltText:  "System Overview",
		Position: docxupdater.PositionBeforeText,
		Anchor:   "Figure 1:",
	})
	if err != nil {
		t.Fatalf("Failed to insert image before text: %v", err)
	}

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the output file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Verify document contains both images
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Introduction Diagram") {
		t.Error("First image not found")
	}
	if !strings.Contains(docXML, "System Overview") {
		t.Error("Second image not found")
	}

	t.Logf("Successfully created document with anchored images at: %s", outputPath)
}

// TestInsertImageVariousSizes tests different sizing options
func TestInsertImageVariousSizes(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_various_sizes_test.docx")

	// Create test image with known dimensions
	tempDir := t.TempDir()
	testImagePath := filepath.Join(tempDir, "test.png")
	createTestImageForTemplate(t, testImagePath, 1200, 800) // 3:2 aspect ratio

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Test 1: Width only (proportional height)
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Test 1: Width=600px (Height should be 400px proportionally)",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
		Bold:     true,
	})
	u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Width:    600,
		AltText:  "Width Only - 600px",
		Position: docxupdater.PositionEnd,
	})

	// Test 2: Height only (proportional width)
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Test 2: Height=300px (Width should be 450px proportionally)",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
		Bold:     true,
	})
	u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Height:   300,
		AltText:  "Height Only - 300px",
		Position: docxupdater.PositionEnd,
	})

	// Test 3: Both dimensions (exact size)
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Test 3: Width=500px, Height=500px (square, may distort)",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
		Bold:     true,
	})
	u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Width:    500,
		Height:   500,
		AltText:  "Both Dimensions - 500x500px",
		Position: docxupdater.PositionEnd,
	})

	// Test 4: No dimensions (actual size)
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Test 4: No dimensions specified (actual size: 1200x800px)",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
		Bold:     true,
	})
	u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		AltText:  "Actual Size - 1200x800px",
		Position: docxupdater.PositionEnd,
	})

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify all images are present
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	images := []string{"Width Only", "Height Only", "Both Dimensions", "Actual Size"}
	for _, img := range images {
		if !strings.Contains(docXML, img) {
			t.Errorf("Image not found: %s", img)
		}
	}

	t.Logf("Successfully created document with various image sizes at: %s", outputPath)
}

// createTestImageForTemplate creates a colorful test image with gradient
func createTestImageForTemplate(t *testing.T, path string, width, height int) {
	t.Helper()

	// Create a gradient image
	img := image.NewRGBA(image.Rect(0, 0, width, height))
	for y := range height {
		for x := range width {
			// Create a nice gradient pattern
			r := uint8((x * 255) / width)
			g := uint8((y * 255) / height)
			b := uint8(((x + y) * 255) / (width + height))
			img.Set(x, y, color.RGBA{r, g, b, 255})
		}
	}

	// Add a border
	borderColor := color.RGBA{0, 0, 0, 255}
	for x := range width {
		img.Set(x, 0, borderColor)
		img.Set(x, height-1, borderColor)
	}
	for y := range height {
		img.Set(0, y, borderColor)
		img.Set(width-1, y, borderColor)
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

// TestInsertImageWithCaption tests image insertion with auto-numbered captions
func TestInsertImageWithCaption(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_captioned_images_test.docx")

	// Create test images
	tempDir := t.TempDir()
	image1Path := filepath.Join(tempDir, "chart.png")
	image2Path := filepath.Join(tempDir, "diagram.png")
	createTestImageForTemplate(t, image1Path, 800, 600)
	createTestImageForTemplate(t, image2Path, 1024, 768)

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Add heading
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Image Gallery with Captions",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionEnd,
	})

	// Insert first image with caption after (default for figures)
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image1Path,
		Width:    500,
		AltText:  "Sales Chart",
		Position: docxupdater.PositionEnd,
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "Monthly Sales Performance",
			AutoNumber:  true,
			Position:    docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		t.Fatalf("Failed to insert first image with caption: %v", err)
	}

	// Add some text
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "The chart above shows significant growth in Q1.",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})

	// Insert second image with caption before
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image2Path,
		Height:   400,
		AltText:  "System Diagram",
		Position: docxupdater.PositionEnd,
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "System Architecture Overview",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore,
		},
	})
	if err != nil {
		t.Fatalf("Failed to insert second image with caption: %v", err)
	}

	// Insert third image with centered caption
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     image1Path,
		Width:    450,
		AltText:  "Process Flow",
		Position: docxupdater.PositionEnd,
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "End-to-End Process Flow",
			AutoNumber:  true,
			Position:    docxupdater.CaptionAfter,
			Alignment:   docxupdater.CellAlignCenter,
		},
	})
	if err != nil {
		t.Fatalf("Failed to insert third image with caption: %v", err)
	}

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the output file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Verify document contains captions
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Monthly Sales Performance") {
		t.Error("First caption not found")
	}
	if !strings.Contains(docXML, "System Architecture Overview") {
		t.Error("Second caption not found")
	}
	if !strings.Contains(docXML, "End-to-End Process Flow") {
		t.Error("Third caption not found")
	}

	// Verify SEQ fields for auto-numbering
	if !strings.Contains(docXML, "SEQ Figure") {
		t.Error("SEQ field not found for auto-numbering")
	}

	t.Logf("Successfully created document with captioned images at: %s", outputPath)
}

// TestInsertImageWithManualCaption tests image with manual caption numbering
func TestInsertImageWithManualCaption(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_manual_caption_test.docx")

	// Create test image
	tempDir := t.TempDir()
	testImagePath := filepath.Join(tempDir, "photo.png")
	createTestImageForTemplate(t, testImagePath, 640, 480)

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Insert image with manual caption numbering
	err = u.InsertImage(docxupdater.ImageOptions{
		Path:     testImagePath,
		Width:    400,
		AltText:  "Product Photo",
		Position: docxupdater.PositionEnd,
		Caption: &docxupdater.CaptionOptions{
			Type:         docxupdater.CaptionFigure,
			Description:  "Product XYZ-100",
			AutoNumber:   false,
			ManualNumber: 5, // Start at Figure 5
			Position:     docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		t.Fatalf("Failed to insert image with manual caption: %v", err)
	}

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify document contains the manual caption
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Product XYZ-100") {
		t.Error("Caption description not found")
	}

	t.Logf("Successfully created document with manual caption at: %s", outputPath)
}

// TestInsertBreaksWithRealTemplate tests page and section breaks with the real template
func TestInsertBreaksWithRealTemplate(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	outputDir := "../outputs"
	outputPath := filepath.Join(outputDir, "template_with_breaks_test.docx")

	// Verify template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Add Chapter 1
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Chapter 1: Introduction",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionEnd,
	})

	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This is the first chapter with introductory content.",
		Position: docxupdater.PositionEnd,
	})

	// Page break after Chapter 1
	err = u.InsertPageBreak(docxupdater.BreakOptions{
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert page break: %v", err)
	}

	// Add Chapter 2
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Chapter 2: Methods",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionEnd,
	})

	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This chapter describes the methodology.",
		Position: docxupdater.PositionEnd,
	})

	// Section break (next page) before Chapter 3
	err = u.InsertSectionBreak(docxupdater.BreakOptions{
		Position:    docxupdater.PositionEnd,
		SectionType: docxupdater.SectionBreakNextPage,
	})
	if err != nil {
		t.Fatalf("Failed to insert section break: %v", err)
	}

	// Add Chapter 3
	u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Chapter 3: Results",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionEnd,
	})

	// Save the document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the output file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Verify document contains breaks
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Error("Page break not found")
	}
	if !strings.Contains(docXML, "<w:sectPr>") {
		t.Error("Section break not found")
	}

	t.Logf("Successfully created document with breaks at: %s", outputPath)
}

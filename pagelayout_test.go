package docxupdater

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestSetPageLayoutLetterPortrait(t *testing.T) {
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

	// Set Letter portrait layout
	err = u.SetPageLayout(*PageLayoutLetterPortrait())
	if err != nil {
		t.Fatalf("SetPageLayout failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the page layout was set
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Check page size (Letter: 12240 x 15840)
	if !strings.Contains(docXML, `w:w="12240"`) {
		t.Error("Page width not set to Letter width (12240)")
	}
	if !strings.Contains(docXML, `w:h="15840"`) {
		t.Error("Page height not set to Letter height (15840)")
	}

	// Check margins (1440 = 1 inch)
	if !strings.Contains(docXML, `w:top="1440"`) {
		t.Error("Top margin not set to 1 inch")
	}
	if !strings.Contains(docXML, `w:left="1440"`) {
		t.Error("Left margin not set to 1 inch")
	}
}

func TestSetPageLayoutA4Landscape(t *testing.T) {
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

	// Set A4 landscape layout
	err = u.SetPageLayout(*PageLayoutA4Landscape())
	if err != nil {
		t.Fatalf("SetPageLayout failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the page layout was set
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Check landscape orientation
	if !strings.Contains(docXML, `w:orient="landscape"`) {
		t.Error("Orientation not set to landscape")
	}

	// For A4 landscape, width and height are swapped
	// A4 portrait: 11906 x 16838
	// A4 landscape: 16838 x 11906
	if !strings.Contains(docXML, `w:w="16838"`) {
		t.Error("Page width not set to A4 landscape width (16838)")
	}
	if !strings.Contains(docXML, `w:h="11906"`) {
		t.Error("Page height not set to A4 landscape height (11906)")
	}
}

func TestSetPageLayoutCustom(t *testing.T) {
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

	// Set custom layout with narrow margins
	customLayout := PageLayoutOptions{
		PageWidth:    PageWidthLetter,
		PageHeight:   PageHeightLetter,
		Orientation:  OrientationPortrait,
		MarginTop:    MarginNarrow, // 0.5"
		MarginRight:  MarginNarrow, // 0.5"
		MarginBottom: MarginNarrow, // 0.5"
		MarginLeft:   MarginWide,   // 1.5"
		MarginHeader: MarginHeaderFooterDefault,
		MarginFooter: MarginHeaderFooterDefault,
		MarginGutter: 0,
	}

	err = u.SetPageLayout(customLayout)
	if err != nil {
		t.Fatalf("SetPageLayout failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify custom margins
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	if !strings.Contains(docXML, `w:top="720"`) {
		t.Error("Top margin not set to narrow (720)")
	}
	if !strings.Contains(docXML, `w:left="2160"`) {
		t.Error("Left margin not set to wide (2160)")
	}
}

func TestSectionBreakWithPageLayout(t *testing.T) {
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

	// Add some content
	err = u.AddText("Portrait section", PositionEnd)
	if err != nil {
		t.Fatalf("AddText failed: %v", err)
	}

	// Insert section break with landscape layout
	err = u.InsertSectionBreak(BreakOptions{
		Position:    PositionEnd,
		SectionType: SectionBreakNextPage,
		PageLayout:  PageLayoutLetterLandscape(),
	})
	if err != nil {
		t.Fatalf("InsertSectionBreak failed: %v", err)
	}

	// Add content in landscape section
	err = u.AddText("Landscape section", PositionEnd)
	if err != nil {
		t.Fatalf("AddText failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the section break with landscape layout
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Should contain landscape orientation
	if !strings.Contains(docXML, `w:orient="landscape"`) {
		t.Error("Landscape orientation not found in section break")
	}

	// Should have Letter dimensions (swapped for landscape)
	if !strings.Contains(docXML, `w:w="15840"`) {
		t.Error("Landscape width not correct")
	}
}

func TestPageLayoutHelpers(t *testing.T) {
	// Test that helper functions return correct values

	// Letter Portrait
	letterPortrait := PageLayoutLetterPortrait()
	if letterPortrait.PageWidth != PageWidthLetter {
		t.Errorf("Letter portrait width = %d, want %d", letterPortrait.PageWidth, PageWidthLetter)
	}
	if letterPortrait.PageHeight != PageHeightLetter {
		t.Errorf("Letter portrait height = %d, want %d", letterPortrait.PageHeight, PageHeightLetter)
	}
	if letterPortrait.Orientation != OrientationPortrait {
		t.Errorf("Letter portrait orientation = %s, want %s", letterPortrait.Orientation, OrientationPortrait)
	}

	// Letter Landscape - dimensions should be swapped
	letterLandscape := PageLayoutLetterLandscape()
	if letterLandscape.PageWidth != PageHeightLetter {
		t.Errorf("Letter landscape width = %d, want %d (height of portrait)", letterLandscape.PageWidth, PageHeightLetter)
	}
	if letterLandscape.PageHeight != PageWidthLetter {
		t.Errorf("Letter landscape height = %d, want %d (width of portrait)", letterLandscape.PageHeight, PageWidthLetter)
	}
	if letterLandscape.Orientation != OrientationLandscape {
		t.Errorf("Letter landscape orientation = %s, want %s", letterLandscape.Orientation, OrientationLandscape)
	}

	// A4 Portrait
	a4Portrait := PageLayoutA4Portrait()
	if a4Portrait.PageWidth != PageWidthA4 {
		t.Errorf("A4 portrait width = %d, want %d", a4Portrait.PageWidth, PageWidthA4)
	}
	if a4Portrait.PageHeight != PageHeightA4 {
		t.Errorf("A4 portrait height = %d, want %d", a4Portrait.PageHeight, PageHeightA4)
	}

	// A4 Landscape
	a4Landscape := PageLayoutA4Landscape()
	if a4Landscape.PageWidth != PageHeightA4 {
		t.Errorf("A4 landscape width = %d, want %d", a4Landscape.PageWidth, PageHeightA4)
	}
	if a4Landscape.PageHeight != PageWidthA4 {
		t.Errorf("A4 landscape height = %d, want %d", a4Landscape.PageHeight, PageWidthA4)
	}

	// Legal Portrait
	legalPortrait := PageLayoutLegalPortrait()
	if legalPortrait.PageWidth != PageWidthLegal {
		t.Errorf("Legal portrait width = %d, want %d", legalPortrait.PageWidth, PageWidthLegal)
	}
	if legalPortrait.PageHeight != PageHeightLegal {
		t.Errorf("Legal portrait height = %d, want %d", legalPortrait.PageHeight, PageHeightLegal)
	}

	// All default margins should be 1440 (1 inch)
	layouts := []*PageLayoutOptions{letterPortrait, letterLandscape, a4Portrait, a4Landscape, legalPortrait}
	for i, layout := range layouts {
		if layout.MarginTop != MarginDefault {
			t.Errorf("Layout %d: top margin = %d, want %d", i, layout.MarginTop, MarginDefault)
		}
		if layout.MarginRight != MarginDefault {
			t.Errorf("Layout %d: right margin = %d, want %d", i, layout.MarginRight, MarginDefault)
		}
		if layout.MarginBottom != MarginDefault {
			t.Errorf("Layout %d: bottom margin = %d, want %d", i, layout.MarginBottom, MarginDefault)
		}
		if layout.MarginLeft != MarginDefault {
			t.Errorf("Layout %d: left margin = %d, want %d", i, layout.MarginLeft, MarginDefault)
		}
	}
}

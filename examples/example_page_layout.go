package main

import (
	"fmt"
	"log"

	docx "github.com/falcomza/docx-update"
)

func main() {
	templatePath := "../templates/template.docx" // Use your template path
	outputPath := "../outputs/page_layout_demo.docx"

	// Open the template document
	updater, err := docx.New(templatePath)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Example 1: Set page layout for the entire document
	fmt.Println("Setting Letter portrait layout...")
	err = updater.SetPageLayout(*docx.PageLayoutLetterPortrait())
	if err != nil {
		log.Fatalf("Failed to set page layout: %v", err)
	}

	// Add some content in portrait
	err = updater.AddText("This section is in Letter portrait orientation (8.5\" x 11\")", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	// Example 2: Insert section break with landscape orientation
	fmt.Println("Inserting section break with landscape layout...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutLetterLandscape(),
	})
	if err != nil {
		log.Fatalf("Failed to insert section break: %v", err)
	}

	// Add content in landscape
	err = updater.AddText("This section is in Letter landscape orientation (11\" x 8.5\")", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	// Example 3: Insert another section break with A4 paper size
	fmt.Println("Inserting section break with A4 landscape...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutA4Landscape(),
	})
	if err != nil {
		log.Fatalf("Failed to insert section break: %v", err)
	}

	// Add content in A4 landscape
	err = updater.AddText("This section is in A4 landscape orientation (297mm x 210mm)", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	// Example 4: Custom page layout with narrow margins
	fmt.Println("Inserting section break with custom layout...")
	customLayout := &docx.PageLayoutOptions{
		PageWidth:    docx.PageWidthLegal, // Legal paper: 8.5" x 14"
		PageHeight:   docx.PageHeightLegal,
		Orientation:  docx.OrientationPortrait,
		MarginTop:    docx.MarginNarrow,              // 0.5"
		MarginRight:  docx.MarginNarrow,              // 0.5"
		MarginBottom: docx.MarginNarrow,              // 0.5"
		MarginLeft:   docx.MarginNarrow,              // 0.5"
		MarginHeader: docx.MarginHeaderFooterDefault, // 0.5"
		MarginFooter: docx.MarginHeaderFooterDefault, // 0.5"
		MarginGutter: 0,
	}

	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  customLayout,
	})
	if err != nil {
		log.Fatalf("Failed to insert section break: %v", err)
	}

	// Add content in custom layout
	err = updater.AddText("This section uses Legal paper (8.5\" x 14\") with narrow margins (0.5\" all around)", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	// Example 5: Wide margins for professional documents
	fmt.Println("Inserting section break with wide margins...")
	wideMargins := &docx.PageLayoutOptions{
		PageWidth:    docx.PageWidthLetter,
		PageHeight:   docx.PageHeightLetter,
		Orientation:  docx.OrientationPortrait,
		MarginTop:    docx.MarginWide,    // 1.5"
		MarginRight:  docx.MarginDefault, // 1"
		MarginBottom: docx.MarginWide,    // 1.5"
		MarginLeft:   docx.MarginWide,    // 1.5" (for binding)
		MarginHeader: docx.MarginHeaderFooterDefault,
		MarginFooter: docx.MarginHeaderFooterDefault,
		MarginGutter: 0,
	}

	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  wideMargins,
	})
	if err != nil {
		log.Fatalf("Failed to insert section break: %v", err)
	}

	// Add content with wide margins
	err = updater.AddText("This section has wide margins (1.5\" top/bottom/left, 1\" right) suitable for binding", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	// Save the document
	fmt.Println("Saving document...")
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Println("Document saved successfully to:", outputPath)
	fmt.Println("\nPage Layout Summary:")
	fmt.Println("- Section 1: Letter portrait with default margins")
	fmt.Println("- Section 2: Letter landscape with default margins")
	fmt.Println("- Section 3: A4 landscape with default margins")
	fmt.Println("- Section 4: Legal portrait with narrow margins")
	fmt.Println("- Section 5: Letter portrait with wide margins")
}

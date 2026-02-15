package main

import (
	"fmt"
	"log"

	docxupdater "github.com/falcomza/docx-update/src"
)

func main() {
	// Open the template document
	updater, err := docxupdater.New("./templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add title
	err = updater.AddHeading(1, "Table Row Height Configuration Examples", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}

	// Example 1: Auto height (default)
	err = updater.AddHeading(2, "1. Auto Height (Default - Fits Content)", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Description"},
			{Title: "Notes"},
		},
		Rows: [][]string{
			{"Short text", "Auto"},
			{"This is a longer text that might wrap to multiple lines", "Auto height expands"},
			{"Text", "Default"},
		},
		HeaderBold:        true,
		HeaderBackground:  "4472C4",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "E7E6E6",
		BorderStyle:       docxupdater.BorderSingle,
		// RowHeightRule defaults to RowHeightAuto
	})
	if err != nil {
		log.Fatalf("Failed to insert auto height table: %v", err)
	}

	err = updater.AddText("↑ Rows automatically adjust height to fit content (default behavior).", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 2: Exact height
	err = updater.AddHeading(2, "2. Exact Height (Fixed - No Growth)", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Item"},
			{Title: "Specification"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"Height", "0.5 inch", "720 twips"},
			{"Mode", "Exact", "Fixed"},
			{"Growth", "No expansion", "Locked"},
		},
		HeaderRowHeight:   720, // 0.5 inch (1440 twips = 1 inch)
		HeaderHeightRule:  docxupdater.RowHeightExact,
		RowHeight:         720, // 0.5 inch
		RowHeightRule:     docxupdater.RowHeightExact,
		HeaderBold:        true,
		HeaderBackground:  "70AD47",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "E2EFD9",
		BorderStyle:       docxupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert exact height table: %v", err)
	}

	err = updater.AddText("↑ All rows have exact height of 0.5 inch (720 twips). Content is clipped if too large.", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 3: AtLeast height (minimum with growth)
	err = updater.AddHeading(2, "3. Minimum Height (At Least - Can Grow)", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Content"},
			{Title: "Behavior"},
		},
		Rows: [][]string{
			{"Short", "Uses minimum height"},
			{"This is much longer content that will cause the row to expand beyond the minimum height setting", "Grows as needed"},
			{"Medium length text here", "Adjusts"},
		},
		RowHeight:         500, // Minimum 500 twips
		RowHeightRule:     docxupdater.RowHeightAtLeast,
		HeaderBold:        true,
		HeaderBackground:  "2E75B5",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "DEEBF7",
		BorderStyle:       docxupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert atLeast height table: %v", err)
	}

	err = updater.AddText("↑ Rows have minimum height of 500 twips but can grow if content requires it.", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 4: Different header and row heights
	err = updater.AddHeading(2, "4. Different Header and Row Heights", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Quarter", Alignment: docxupdater.CellAlignCenter},
			{Title: "Revenue", Alignment: docxupdater.CellAlignRight},
			{Title: "Growth", Alignment: docxupdater.CellAlignRight},
		},
		Rows: [][]string{
			{"Q1 2026", "$250,000", "+12%"},
			{"Q2 2026", "$280,000", "+15%"},
			{"Q3 2026", "$310,000", "+18%"},
			{"Q4 2026", "$340,000", "+21%"},
		},
		HeaderRowHeight:   900, // Taller header (0.625 inch)
		HeaderHeightRule:  docxupdater.RowHeightExact,
		RowHeight:         450, // Shorter data rows (0.3125 inch)
		RowHeightRule:     docxupdater.RowHeightExact,
		HeaderBold:        true,
		HeaderBackground:  "C65911",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "FCE4D6",
		BorderStyle:       docxupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert mixed height table: %v", err)
	}

	err = updater.AddText("↑ Header row is 900 twips (0.625\") tall, data rows are 450 twips (0.3125\") tall.", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 5: Tall rows for better spacing
	err = updater.AddHeading(2, "5. Tall Rows for Visual Impact", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Task", Alignment: docxupdater.CellAlignLeft},
			{Title: "Status", Alignment: docxupdater.CellAlignCenter},
		},
		Rows: [][]string{
			{"Design Review", "✓ Complete"},
			{"Development", "⟳ In Progress"},
			{"Testing", "◯ Pending"},
		},
		HeaderRowHeight:   1080, // 0.75 inch
		HeaderHeightRule:  docxupdater.RowHeightExact,
		RowHeight:         720, // 0.5 inch - spacious rows
		RowHeightRule:     docxupdater.RowHeightExact,
		HeaderBold:        true,
		HeaderBackground:  "7030A0",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "E9D8F4",
		BorderStyle:       docxupdater.BorderSingle,
		TableWidthType:    docxupdater.TableWidthPercentage,
		TableWidth:        3750, // 75% width
		TableAlignment:    docxupdater.AlignCenter,
	})
	if err != nil {
		log.Fatalf("Failed to insert tall rows table: %v", err)
	}

	err = updater.AddText("↑ Larger row heights (0.5-0.75 inch) create more spacious, easier-to-read tables.", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Add summary section
	err = updater.AddHeading(2, "Row Height Summary", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add summary heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Rule", Alignment: docxupdater.CellAlignLeft},
			{Title: "Behavior", Alignment: docxupdater.CellAlignLeft},
			{Title: "Use Case", Alignment: docxupdater.CellAlignLeft},
		},
		Rows: [][]string{
			{"RowHeightAuto", "Fits content automatically", "Default - variable content"},
			{"RowHeightExact", "Fixed height, no growth", "Uniform appearance, forms"},
			{"RowHeightAtLeast", "Minimum height, can grow", "Consistent minimum, flexible"},
		},
		HeaderBold:        true,
		HeaderBackground:  "44546A",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "D6DCE4",
		BorderStyle:       docxupdater.BorderSingle,
		RowStyle: docxupdater.CellStyle{
			FontSize: 18, // 9pt
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert summary table: %v", err)
	}

	// Add measurement reference
	err = updater.AddHeading(2, "Height Measurements (Twips)", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add measurements heading: %v", err)
	}

	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Inches", Alignment: docxupdater.CellAlignRight},
			{Title: "Twips", Alignment: docxupdater.CellAlignRight},
			{Title: "Common Use", Alignment: docxupdater.CellAlignLeft},
		},
		Rows: [][]string{
			{"0.25\"", "360", "Compact rows"},
			{"0.3125\"", "450", "Standard data rows"},
			{"0.5\"", "720", "Spacious rows"},
			{"0.625\"", "900", "Large header rows"},
			{"0.75\"", "1080", "Extra spacious"},
			{"1.0\"", "1440", "Very tall rows"},
		},
		HeaderBold:        true,
		HeaderBackground:  "203864",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "D9E2F3",
		BorderStyle:       docxupdater.BorderSingle,
		TableWidthType:    docxupdater.TableWidthPercentage,
		TableWidth:        3000, // 60% width
	})
	if err != nil {
		log.Fatalf("Failed to insert measurements table: %v", err)
	}

	err = updater.AddText("Note: 1 inch = 1440 twips. Formula: height_in_twips = height_in_inches × 1440", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add note: %v", err)
	}

	// Save the document
	outputPath := "./outputs/table_row_height_examples.docx"
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Println("✅ SUCCESS!")
	fmt.Printf("📄 Output saved to: %s\n", outputPath)
	fmt.Println("\nCreated examples:")
	fmt.Println("  • Auto height (default - fits content)")
	fmt.Println("  • Exact height (fixed - no growth)")
	fmt.Println("  • Minimum height (at least - can grow)")
	fmt.Println("  • Different header and row heights")
	fmt.Println("  • Tall rows for visual impact")
	fmt.Println("  • Summary and measurement tables")
}

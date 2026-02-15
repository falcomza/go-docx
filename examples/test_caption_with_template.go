package main

import (
	"log"
	"path/filepath"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Use the template file
	templatePath := filepath.Join("templates", "docx_template.docx")
	outputPath := filepath.Join("outputs", "caption_test_output.docx")

	u, err := updater.New(templatePath)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Insert a table with caption
	log.Println("Inserting table with caption...")
	err = u.InsertTable(updater.TableOptions{
		Position: updater.PositionEnd,
		Columns: []updater.ColumnDefinition{
			{Title: "Product"},
			{Title: "Q1"},
			{Title: "Q2"},
			{Title: "Q3"},
			{Title: "Q4"},
		},
		Rows: [][]string{
			{"Product A", "$120K", "$135K", "$128K", "$150K"},
			{"Product B", "$98K", "$105K", "$112K", "$118K"},
			{"Product C", "$85K", "$92K", "$88K", "$95K"},
		},
		HeaderBold:      true,
		HeaderAlignment: updater.CellAlignCenter,
		Caption: &updater.CaptionOptions{
			Type:        updater.CaptionTable,
			Description: "Quarterly Sales Data",
			AutoNumber:  true,
			Position:    updater.CaptionBefore,
			Alignment:   updater.CellAlignCenter,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert table: %v", err)
	}

	// Insert a chart with caption
	log.Println("Inserting chart with caption...")
	err = u.InsertChart(updater.ChartOptions{
		Position:   updater.PositionEnd,
		Title:      "Quarterly Revenue",
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []updater.SeriesData{
			{Name: "2025", Values: []float64{100, 120, 110, 130}},
			{Name: "2026", Values: []float64{110, 130, 125, 145}},
		},
		ShowLegend: true,
		Caption: &updater.CaptionOptions{
			Type:        updater.CaptionFigure,
			Description: "Revenue Trends 2025-2026",
			AutoNumber:  true,
			Position:    updater.CaptionAfter,
			Alignment:   updater.CellAlignCenter,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	// Save the output
	if err := u.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	log.Printf("✓ Successfully created document with captions: %s", outputPath)
	log.Println("✓ Table caption: 'Table 1: Quarterly Sales Data'")
	log.Println("✓ Chart caption: 'Figure 1: Revenue Trends 2025-2026'")
}

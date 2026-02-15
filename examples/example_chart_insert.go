package main

import (
	"log"

	docxupdater "github.com/falcomza/docx-update/src"
)

func main() {
	// Open the input document
	u, err := docxupdater.New("input.docx")
	if err != nil {
		log.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	// Example 1: Basic sales chart
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Quarterly Sales Report 2024",
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []docxupdater.SeriesData{
			{Name: "Revenue", Values: []float64{250000, 280000, 310000, 290000}},
			{Name: "Profit", Values: []float64{50000, 62000, 68000, 64000}},
		},
		ShowLegend: true,
	})
	if err != nil {
		log.Fatalf("Failed to insert sales chart: %v", err)
	}

	// Example 2: Chart with axis titles
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:          docxupdater.PositionEnd,
		Title:             "Website Traffic Analysis",
		CategoryAxisTitle: "Month",
		ValueAxisTitle:    "Visitors (thousands)",
		Categories:        []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun"},
		Series: []docxupdater.SeriesData{
			{Name: "Unique Visitors", Values: []float64{45, 52, 58, 61, 65, 70}},
			{Name: "Page Views", Values: []float64{180, 220, 240, 255, 275, 295}},
		},
		ShowLegend:     true,
		LegendPosition: "b", // Bottom position
	})
	if err != nil {
		log.Fatalf("Failed to insert traffic chart: %v", err)
	}

	// Example 3: Multi-series financial chart
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:          docxupdater.PositionEnd,
		Title:             "Financial Performance",
		CategoryAxisTitle: "Period",
		ValueAxisTitle:    "Amount (USD)",
		Categories:        []string{"Jan", "Feb", "Mar", "Apr"},
		Series: []docxupdater.SeriesData{
			{Name: "Revenue", Values: []float64{100000, 120000, 115000, 130000}},
			{Name: "Costs", Values: []float64{60000, 70000, 65000, 75000}},
			{Name: "Profit", Values: []float64{40000, 50000, 50000, 55000}},
		},
		ShowLegend:     true,
		LegendPosition: "r", // Right position (default)
	})
	if err != nil {
		log.Fatalf("Failed to insert financial chart: %v", err)
	}

	// Example 4: Chart with custom dimensions
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Product Comparison",
		Categories: []string{"Product A", "Product B", "Product C"},
		Series: []docxupdater.SeriesData{
			{Name: "Sales", Values: []float64{150, 200, 175}},
		},
		ShowLegend: false,   // No legend for single series
		Width:      5486400, // ~6 inches
		Height:     3048000, // ~3.3 inches
	})
	if err != nil {
		log.Fatalf("Failed to insert product chart: %v", err)
	}

	// Save the output document
	if err := u.Save("output_with_charts.docx"); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	log.Println("Successfully created document with charts!")
	log.Println("Output saved to: output_with_charts.docx")
}

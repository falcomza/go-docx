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

	// Example 1: Chart with automatic caption (default position: after)
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Quarterly Sales Report 2024",
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []docxupdater.SeriesData{
			{Name: "Revenue", Values: []float64{250000, 280000, 310000, 290000}},
			{Name: "Profit", Values: []float64{50000, 62000, 68000, 64000}},
		},
		ShowLegend: true,
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "Quarterly sales performance showing revenue and profit trends",
			AutoNumber:  true,
			Position:    docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert sales chart with caption: %v", err)
	}

	// Example 2: Chart with caption before the chart (less common for figures)
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
		LegendPosition: "b",
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "Monthly website traffic metrics for H1 2024",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore,
			Alignment:   docxupdater.CellAlignCenter,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert traffic chart with caption: %v", err)
	}

	// Example 3: Table with caption (default position: before for tables)
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Product", Width: 2000},
			{Title: "Q1 Sales", Width: 1500},
			{Title: "Q2 Sales", Width: 1500},
			{Title: "Q3 Sales", Width: 1500},
			{Title: "Q4 Sales", Width: 1500},
		},
		Rows: [][]string{
			{"Product A", "$50,000", "$55,000", "$60,000", "$58,000"},
			{"Product B", "$45,000", "$48,000", "$52,000", "$54,000"},
			{"Product C", "$38,000", "$42,000", "$45,000", "$47,000"},
		},
		HeaderBold:      true,
		HeaderAlignment: docxupdater.CellAlignCenter,
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionTable,
			Description: "Product sales by quarter for 2024",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore, // Default for tables
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert table with caption: %v", err)
	}

	// Example 4: Table with centered caption
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Department"},
			{Title: "Budget"},
			{Title: "Spent"},
			{Title: "Remaining"},
		},
		Rows: [][]string{
			{"Marketing", "$100,000", "$85,000", "$15,000"},
			{"R&D", "$200,000", "$180,000", "$20,000"},
			{"Sales", "$150,000", "$140,000", "$10,000"},
		},
		HeaderBold:       true,
		HeaderAlignment:  docxupdater.CellAlignCenter,
		HeaderBackground: "4472C4",
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionTable,
			Description: "Department budget allocation and expenditure",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore,
			Alignment:   docxupdater.CellAlignCenter,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert budget table with caption: %v", err)
	}

	// Example 5: Using default caption options
	defaultCaption := docxupdater.DefaultCaptionOptions(docxupdater.CaptionFigure)
	defaultCaption.Description = "Growth trend analysis for key metrics"

	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Annual Growth Rate",
		Categories: []string{"2020", "2021", "2022", "2023", "2024"},
		Series: []docxupdater.SeriesData{
			{Name: "Growth %", Values: []float64{5.2, 6.8, 7.1, 6.5, 8.3}},
		},
		ShowLegend: false,
		Caption:    &defaultCaption,
	})
	if err != nil {
		log.Fatalf("Failed to insert growth chart with caption: %v", err)
	}

	// Example 6: Manual numbering (no auto-numbering)
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Market Share Distribution",
		Categories: []string{"Company A", "Company B", "Company C", "Others"},
		Series: []docxupdater.SeriesData{
			{Name: "Market Share", Values: []float64{35, 28, 22, 15}},
		},
		ShowLegend: false,
		Caption: &docxupdater.CaptionOptions{
			Type:         docxupdater.CaptionFigure,
			Description:  "Current market share distribution by company",
			AutoNumber:   false,
			ManualNumber: 99, // Custom number
			Position:     docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert market chart with manual caption: %v", err)
	}

	// Save the output document
	if err := u.Save("output_with_captions.docx"); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	log.Println("Successfully created document with captions!")
	log.Println("Output saved to: output_with_captions.docx")
	log.Println("\nCaptions include:")
	log.Println("- Figure captions with automatic numbering")
	log.Println("- Table captions positioned before tables")
	log.Println("- Custom alignment and positioning options")
	log.Println("- Manual numbering example")
}

package docxupdater

import (
	"fmt"
	"testing"
	"time"
)

// TestComprehensiveDemo demonstrates all features with the template
func TestComprehensiveDemo(t *testing.T) {
	t.Run("ChartUpdate", testChartUpdateDemo)
	t.Run("Properties", testPropertiesDemo)
	t.Run("Bookmarks", testBookmarksDemo)
	t.Run("InsertChart", testInsertChartDemo)
	t.Run("CompleteWorkflow", testCompleteWorkflowDemo)
	t.Run("TextOperations", testTextOperationsDemo)
	t.Run("TableAndImage", testTableAndImageDemo)
}

// testChartUpdateDemo creates a new chart instead of updating existing one
func testChartUpdateDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add heading
	if err := updater.AddHeading(1, "Financial Performance Report", PositionBeginning); err != nil {
		t.Fatalf("failed to add heading: %v", err)
	}

	// Insert chart with quarterly sales data
	if err := updater.InsertChart(ChartOptions{
		Position:          PositionEnd,
		ChartKind:         ChartKindColumn,
		Title:             "Quarterly Financial Performance 2024",
		CategoryAxisTitle: "Fiscal Quarter",
		ValueAxisTitle:    "Amount (USD)",
		Categories:        []string{"Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"},
		Series: []SeriesData{
			{Name: "Revenue", Values: []float64{125000, 145000, 132000, 168000}},
			{Name: "Expenses", Values: []float64{85000, 95000, 88000, 102000}},
			{Name: "Profit", Values: []float64{40000, 50000, 44000, 66000}},
		},
		ShowLegend:     true,
		LegendPosition: "r",
	}); err != nil {
		t.Fatalf("failed to insert chart: %v", err)
	}

	if err := updater.Save("outputs/demo_chart_update.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Chart update demo completed: outputs/demo_chart_update.docx")
}

// testPropertiesDemo sets document properties and adds content
func testPropertiesDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add some content first
	if err := updater.AddHeading(1, "Document Properties Demo", PositionBeginning); err != nil {
		t.Fatalf("failed to add heading: %v", err)
	}

	if err := updater.AddText("This document demonstrates setting document properties including core, app, and custom properties.", PositionEnd); err != nil {
		t.Fatalf("failed to add text: %v", err)
	}

	// Set core properties
	if err := updater.SetCoreProperties(CoreProperties{
		Title:       "Properties Demo Report",
		Subject:     "DOCX Library Properties",
		Creator:     "Test Suite",
		Keywords:    "docx, properties, metadata",
		Description: "Demonstrates property management",
		Category:    "Demo",
	}); err != nil {
		t.Fatalf("failed to set core properties: %v", err)
	}

	// Set app properties
	if err := updater.SetAppProperties(AppProperties{
		Company:     "Test Corp",
		Application: "DOCX Library",
	}); err != nil {
		t.Fatalf("failed to set app properties: %v", err)
	}

	if err := updater.Save("outputs/demo_properties.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Properties demo completed: outputs/demo_properties.docx")
}

// testBookmarksDemo creates bookmarks in the document
func testBookmarksDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Create a bookmark at the beginning
	err = updater.CreateBookmark("doc_start", BookmarkOptions{
		Position: PositionBeginning,
		Hidden:   true,
	})
	if err != nil {
		t.Fatalf("failed to create bookmark: %v", err)
	}

	// Create a bookmark with text
	err = updater.CreateBookmarkWithText("section_intro", "Important Section", BookmarkOptions{
		Position: PositionEnd,
		Style:    StyleHeading2,
	})
	if err != nil {
		t.Fatalf("failed to create bookmark with text: %v", err)
	}

	if err := updater.Save("outputs/demo_bookmarks.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Bookmarks demo completed: outputs/demo_bookmarks.docx")
}

// testInsertChartDemo creates a new chart dynamically
func testInsertChartDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Insert a new column chart
	err = updater.InsertChart(ChartOptions{
		Position:          PositionEnd,
		ChartKind:         ChartKindColumn,
		Title:             "Monthly Sales Trends",
		CategoryAxisTitle: "Month",
		ValueAxisTitle:    "Sales (Units)",
		Categories:        []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun"},
		Series: []SeriesData{
			{Name: "Product A", Values: []float64{120, 135, 150, 145, 160, 175}},
			{Name: "Product B", Values: []float64{80, 85, 95, 100, 105, 110}},
		},
		ShowLegend:     true,
		LegendPosition: "r",
		Caption: &CaptionOptions{
			Type:        CaptionFigure,
			Description: "Monthly sales performance by product",
			AutoNumber:  true,
		},
	})
	if err != nil {
		t.Fatalf("failed to insert chart: %v", err)
	}

	if err := updater.Save("outputs/demo_insert_chart.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Insert chart demo completed: outputs/demo_insert_chart.docx")
}

// testCompleteWorkflowDemo demonstrates a complete document generation workflow
func testCompleteWorkflowDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// 1. Set properties
	if err := updater.SetCoreProperties(CoreProperties{
		Title:   "Complete Workflow Demo",
		Creator: "Test System",
		Subject: "Full Feature Demonstration",
	}); err != nil {
		t.Fatalf("failed to set properties: %v", err)
	}

	// 2. Add heading
	if err := updater.AddHeading(1, "Executive Summary", PositionBeginning); err != nil {
		t.Fatalf("failed to add heading: %v", err)
	}

	// 3. Add content paragraph
	if err := updater.AddText("This document demonstrates a complete workflow using the DOCX Update library. "+
		"It includes charts, tables, and various formatting options.", PositionEnd); err != nil {
		t.Fatalf("failed to add text: %v", err)
	}

	// 4. Insert a table
	if err := updater.InsertTable(TableOptions{
		Position: PositionEnd,
		Columns: []ColumnDefinition{
			{Title: "Metric", Width: 2500, Bold: true},
			{Title: "2023", Width: 1500},
			{Title: "2024", Width: 1500},
			{Title: "Change", Width: 1500},
		},
		Rows: [][]string{
			{"Revenue", "$1.2M", "$1.5M", "+25%"},
			{"Customers", "1,200", "1,650", "+37.5%"},
			{"Market Share", "15%", "18%", "+3pp"},
		},
		HeaderBackground: "4472C4",
		HeaderBold:       true,
		TableStyle:       TableStyleProfessional,
		Caption: &CaptionOptions{
			Type:        CaptionTable,
			Description: "Key performance metrics comparison",
			AutoNumber:  true,
		},
	}); err != nil {
		t.Fatalf("failed to insert table: %v", err)
	}

	// 5. Add page break
	if err := updater.InsertPageBreak(BreakOptions{
		Position: PositionEnd,
	}); err != nil {
		t.Fatalf("failed to insert page break: %v", err)
	}

	// 6. Insert new chart
	if err := updater.InsertChart(ChartOptions{
		Position:   PositionEnd,
		ChartKind:  ChartKindLine,
		Title:      "Customer Acquisition Trend",
		Categories: []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun"},
		Series: []SeriesData{
			{Name: "New Customers", Values: []float64{45, 52, 58, 64, 70, 78}},
		},
		ShowLegend: true,
	}); err != nil {
		t.Fatalf("failed to insert chart: %v", err)
	}

	// 7. Add footer
	if err := updater.SetFooter(HeaderFooterContent{
		LeftText:         "Confidential",
		CenterText:       "Complete Workflow Demo",
		PageNumber:       true,
		PageNumberFormat: "Page X of Y",
	}, FooterOptions{
		Type: FooterDefault,
	}); err != nil {
		t.Fatalf("failed to set footer: %v", err)
	}

	if err := updater.Save("outputs/demo_complete_workflow.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Complete workflow demo completed: outputs/demo_complete_workflow.docx")
}

// testTextOperationsDemo tests text search and replace
func testTextOperationsDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add some text with placeholders
	updater.AddText("Company: {{COMPANY_NAME}}", PositionEnd)
	updater.AddText("Date: {{DATE}}", PositionEnd)
	updater.AddText("Report Type: {{TYPE}}", PositionEnd)

	// Replace placeholders
	updater.ReplaceText("{{COMPANY_NAME}}", "Acme Corporation", ReplaceOptions{
		InParagraphs: true,
	})
	updater.ReplaceText("{{DATE}}", time.Now().Format("January 2, 2006"), ReplaceOptions{
		InParagraphs: true,
	})
	updater.ReplaceText("{{TYPE}}", "Quarterly Financial Report", ReplaceOptions{
		InParagraphs: true,
	})

	// Insert hyperlink
	updater.InsertHyperlink("Visit our website", "https://example.com", HyperlinkOptions{
		Position:  PositionEnd,
		Color:     "0563C1",
		Underline: true,
		Tooltip:   "Click to open website",
	})

	if err := updater.Save("outputs/demo_text_operations.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Text operations demo completed: outputs/demo_text_operations.docx")
}

// testTableAndImageDemo tests table insertion
func testTableAndImageDemo(t *testing.T) {
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Fatalf("failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add heading
	updater.AddHeading(1, "Data Analysis Report", PositionBeginning)

	// Insert a complex table
	updater.InsertTable(TableOptions{
		Position: PositionEnd,
		Columns: []ColumnDefinition{
			{Title: "Region", Width: 2000, Bold: true},
			{Title: "Q1", Width: 1200},
			{Title: "Q2", Width: 1200},
			{Title: "Q3", Width: 1200},
			{Title: "Q4", Width: 1200},
			{Title: "Total", Width: 1500, Bold: true},
		},
		Rows: [][]string{
			{"North America", "$250K", "$275K", "$290K", "$310K", "$1,125K"},
			{"Europe", "$180K", "$195K", "$205K", "$220K", "$800K"},
			{"Asia Pacific", "$150K", "$165K", "$180K", "$200K", "$695K"},
			{"Latin America", "$90K", "$95K", "$100K", "$105K", "$390K"},
			{"Total", "$670K", "$730K", "$775K", "$835K", "$3,010K"},
		},
		HeaderBackground:  "2E75B6",
		HeaderBold:        true,
		AlternateRowColor: "D9E2F3",
		TableStyle:        TableStyleProfessional,
		RepeatHeader:      true,
		Caption: &CaptionOptions{
			Type:        CaptionTable,
			Description: "Regional sales performance by quarter",
			AutoNumber:  true,
		},
	})

	// Add section break
	updater.InsertSectionBreak(BreakOptions{
		Position:    PositionEnd,
		SectionType: SectionBreakNextPage,
	})

	// Add narrative text
	updater.AddHeading(2, "Analysis Summary", PositionEnd)
	updater.AddText("The data shows consistent growth across all regions throughout the year. "+
		"North America continues to be the strongest market, while Asia Pacific shows the "+
		"highest growth rate.", PositionEnd)

	if err := updater.Save("outputs/demo_table_analysis.docx"); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	t.Log("✓ Table and analysis demo completed: outputs/demo_table_analysis.docx")
}

// RunComprehensiveDemo can be called directly to generate all demo files
func RunComprehensiveDemo() {
	fmt.Println("=== Running Comprehensive Demo Suite ===")
	fmt.Println()

	tests := []struct {
		name string
		fn   func(*testing.T)
	}{
		{"Chart Update", testChartUpdateDemo},
		{"Properties", testPropertiesDemo},
		{"Bookmarks", testBookmarksDemo},
		{"Insert Chart", testInsertChartDemo},
		{"Complete Workflow", testCompleteWorkflowDemo},
		{"Text Operations", testTextOperationsDemo},
		{"Table Analysis", testTableAndImageDemo},
	}

	for _, tt := range tests {
		fmt.Printf("Running %s...\n", tt.name)
		t := &testing.T{}
		tt.fn(t)
		if t.Failed() {
			fmt.Printf("  ✗ %s failed\n", tt.name)
		} else {
			fmt.Printf("  ✓ %s completed\n", tt.name)
		}
	}

	fmt.Println()
	fmt.Println("=== All demos completed ===")
	fmt.Println("Check the ./outputs folder for generated documents:")
	fmt.Println("  - demo_chart_update.docx")
	fmt.Println("  - demo_properties.docx")
	fmt.Println("  - demo_bookmarks.docx")
	fmt.Println("  - demo_insert_chart.docx")
	fmt.Println("  - demo_complete_workflow.docx")
	fmt.Println("  - demo_text_operations.docx")
	fmt.Println("  - demo_table_analysis.docx")
}

package docxupdater_test

import (
	"archive/zip"
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

func TestInsertBasicChart(t *testing.T) {
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

	// Create a basic column chart
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Sales Report",
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []docxupdater.SeriesData{
			{Name: "Revenue", Values: []float64{100, 150, 120, 180}},
			{Name: "Profit", Values: []float64{20, 30, 25, 40}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// List all files in output
	entries := listZipEntries(t, outputPath)
	t.Logf("Files in output: %v", entries)

	// Find which chart was created
	var chartFile string
	for _, entry := range entries {
		if strings.HasPrefix(entry, "word/charts/chart") && strings.HasSuffix(entry, ".xml") {
			if strings.Contains(entry, ".rels") {
				continue
			}
			chartFile = entry
			t.Logf("Found chart file: %s", chartFile)
		}
	}

	if chartFile == "" {
		t.Fatal("No chart file found in output")
	}

	// Verify chart was created
	chartXML := readZipEntry(t, outputPath, chartFile)
	t.Logf("Chart XML length: %d", len(chartXML))
	t.Logf("Chart XML (first 1000 chars): %s", chartXML[:min(1000, len(chartXML))])
	if !strings.Contains(chartXML, "Sales Report") {
		t.Error("Chart title not found in chart XML")
	}
	if !strings.Contains(chartXML, "Revenue") || !strings.Contains(chartXML, "Profit") {
		t.Error("Series names not found in chart XML")
	}

	// Verify embedded workbook was created
	workbookExists := false
	for _, entry := range entries {
		if strings.Contains(entry, "Microsoft_Excel_Worksheet") {
			workbookExists = true
			break
		}
	}
	if !workbookExists {
		t.Error("Embedded workbook not found")
	}

	// Verify chart drawing in document
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<c:chart") {
		t.Error("Chart drawing not found in document.xml")
	}
}

func TestInsertChartWithAxisTitles(t *testing.T) {
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

	err = u.InsertChart(docxupdater.ChartOptions{
		Position:          docxupdater.PositionEnd,
		Title:             "Performance Metrics",
		CategoryAxisTitle: "Time Period",
		ValueAxisTitle:    "Value (USD)",
		Categories:        []string{"Jan", "Feb", "Mar"},
		Series: []docxupdater.SeriesData{
			{Name: "Sales", Values: []float64{1000, 1200, 1500}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	chartXML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	if !strings.Contains(chartXML, "Time Period") {
		t.Error("Category axis title not found")
	}
	if !strings.Contains(chartXML, "Value (USD)") {
		t.Error("Value axis title not found")
	}
}

func TestInsertMultipleCharts(t *testing.T) {
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

	// Insert first chart
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Chart 1",
		Categories: []string{"A", "B"},
		Series: []docxupdater.SeriesData{
			{Name: "Series1", Values: []float64{10, 20}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart 1 failed: %v", err)
	}

	// Insert second chart
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Chart 2",
		Categories: []string{"X", "Y"},
		Series: []docxupdater.SeriesData{
			{Name: "Series2", Values: []float64{30, 40}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart 2 failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify both charts exist
	chart1XML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	if !strings.Contains(chart1XML, "Chart 1") {
		t.Error("Chart 1 not found")
	}

	chart2XML := readZipEntry(t, outputPath, "word/charts/chart2.xml")
	if !strings.Contains(chart2XML, "Chart 2") {
		t.Error("Chart 2 not found")
	}
}

func TestInsertChartInvalidData(t *testing.T) {
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

	// Test empty categories
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Categories: []string{},
		Series: []docxupdater.SeriesData{
			{Name: "Test", Values: []float64{}},
		},
	})
	if err == nil {
		t.Error("Expected error for empty categories")
	}

	// Test empty series
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Categories: []string{"A", "B"},
		Series:     []docxupdater.SeriesData{},
	})
	if err == nil {
		t.Error("Expected error for empty series")
	}

	// Test mismatched values length
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Categories: []string{"A", "B", "C"},
		Series: []docxupdater.SeriesData{
			{Name: "Test", Values: []float64{1, 2}}, // Only 2 values, but 3 categories
		},
	})
	if err == nil {
		t.Error("Expected error for mismatched values length")
	}
}

func TestInsertChartMultipleSeries(t *testing.T) {
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

	err = u.InsertChart(docxupdater.ChartOptions{
		Position:          docxupdater.PositionEnd,
		Title:             "Sales vs Costs",
		CategoryAxisTitle: "Month",
		ValueAxisTitle:    "Amount",
		Categories:        []string{"Jan", "Feb", "Mar", "Apr"},
		Series: []docxupdater.SeriesData{
			{Name: "Revenue", Values: []float64{1000, 1200, 1100, 1300}},
			{Name: "Costs", Values: []float64{600, 700, 650, 750}},
			{Name: "Profit", Values: []float64{400, 500, 450, 550}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	chartXML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	if !strings.Contains(chartXML, "Revenue") {
		t.Error("Revenue series not found")
	}
	if !strings.Contains(chartXML, "Costs") {
		t.Error("Costs series not found")
	}
	if !strings.Contains(chartXML, "Profit") {
		t.Error("Profit series not found")
	}
}

func TestInsertChartNoLegend(t *testing.T) {
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

	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Chart Without Legend",
		Categories: []string{"A", "B"},
		Series: []docxupdater.SeriesData{
			{Name: "Data", Values: []float64{10, 20}},
		},
		ShowLegend: false,
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	chartXML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	// Legend should not be present when ShowLegend is false
	if strings.Contains(chartXML, "<c:legend>") {
		t.Error("Legend found when ShowLegend is false")
	}
}

func TestInsertChartAtBeginning(t *testing.T) {
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

	// Add some text first
	if err := u.AddText("This is after the chart", docxupdater.PositionEnd); err != nil {
		t.Fatalf("AddText failed: %v", err)
	}

	// Insert chart at beginning
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionBeginning,
		Title:      "First Chart",
		Categories: []string{"A", "B"},
		Series: []docxupdater.SeriesData{
			{Name: "Data", Values: []float64{5, 10}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify chart exists
	chartXML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	if !strings.Contains(chartXML, "First Chart") {
		t.Error("Chart not found")
	}
}

// Helper function to list all entries in a zip file
func listZipEntries(t *testing.T, zipPath string) []string {
	t.Helper()

	reader, err := zip.OpenReader(zipPath)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer reader.Close()

	var entries []string
	for _, f := range reader.File {
		entries = append(entries, f.Name)
	}
	return entries
}

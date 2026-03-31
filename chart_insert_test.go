package godocx_test

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

func TestInsertBasicChart(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create a basic column chart
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Title:      "Sales Report",
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []godocx.SeriesOptions{
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
	chartXML, chartFile := findChartXMLContaining(t, outputPath, "Sales Report")
	t.Logf("Chart XML length: %d", len(chartXML))
	t.Logf("Chart XML (first 1000 chars): %s", chartXML[:min(1000, len(chartXML))])
	t.Logf("Verified chart file: %s", chartFile)
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertChart(godocx.ChartOptions{
		Position:          godocx.PositionEnd,
		Title:             "Performance Metrics",
		CategoryAxisTitle: "Time Period",
		ValueAxisTitle:    "Value (USD)",
		Categories:        []string{"Jan", "Feb", "Mar"},
		Series: []godocx.SeriesOptions{
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

	chartXML, _ := findChartXMLContaining(t, outputPath, "Performance Metrics")
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert first chart
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Title:      "Chart 1",
		Categories: []string{"A", "B"},
		Series: []godocx.SeriesOptions{
			{Name: "Series1", Values: []float64{10, 20}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart 1 failed: %v", err)
	}

	// Insert second chart
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Title:      "Chart 2",
		Categories: []string{"X", "Y"},
		Series: []godocx.SeriesOptions{
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

	// Verify both chart titles exist in chart files
	chart1XML, _ := findChartXMLContaining(t, outputPath, "Chart 1")
	if !strings.Contains(chart1XML, "Chart 1") {
		t.Error("Chart 1 not found")
	}

	chart2XML, _ := findChartXMLContaining(t, outputPath, "Chart 2")
	if !strings.Contains(chart2XML, "Chart 2") {
		t.Error("Chart 2 not found")
	}

	assertEmbeddedWorkbookContentTypes(t, outputPath)
}

func TestInsertChartInvalidData(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Test empty categories
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Categories: []string{},
		Series: []godocx.SeriesOptions{
			{Name: "Test", Values: []float64{}},
		},
	})
	if err == nil {
		t.Error("Expected error for empty categories")
	}

	// Test empty series
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Categories: []string{"A", "B"},
		Series:     []godocx.SeriesOptions{},
	})
	if err == nil {
		t.Error("Expected error for empty series")
	}

	// Test mismatched values length
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Categories: []string{"A", "B", "C"},
		Series: []godocx.SeriesOptions{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertChart(godocx.ChartOptions{
		Position:          godocx.PositionEnd,
		Title:             "Sales vs Costs",
		CategoryAxisTitle: "Month",
		ValueAxisTitle:    "Amount",
		Categories:        []string{"Jan", "Feb", "Mar", "Apr"},
		Series: []godocx.SeriesOptions{
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

	chartXML, _ := findChartXMLContaining(t, outputPath, "Sales vs Costs")
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		Title:      "Chart Without Legend",
		Categories: []string{"A", "B"},
		Series: []godocx.SeriesOptions{
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

	chartXML, _ := findChartXMLContaining(t, outputPath, "Chart Without Legend")
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add some text first
	if err := u.AddText("This is after the chart", godocx.PositionEnd); err != nil {
		t.Fatalf("AddText failed: %v", err)
	}

	// Insert chart at beginning
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionBeginning,
		Title:      "First Chart",
		Categories: []string{"A", "B"},
		Series: []godocx.SeriesOptions{
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
	chartXML, _ := findChartXMLContaining(t, outputPath, "First Chart")
	if !strings.Contains(chartXML, "First Chart") {
		t.Error("Chart not found")
	}
}

func findChartXMLContaining(t *testing.T, docxPath string, needle string) (string, string) {
	t.Helper()

	entries := listZipEntries(t, docxPath)
	for _, entry := range entries {
		if !strings.HasPrefix(entry, "word/charts/chart") || !strings.HasSuffix(entry, ".xml") {
			continue
		}
		if strings.Contains(entry, ".rels") {
			continue
		}

		xml := readZipEntry(t, docxPath, entry)
		if strings.Contains(xml, needle) {
			return xml, entry
		}
	}

	t.Fatalf("no chart xml contains %q", needle)
	return "", ""
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

type contentTypes struct {
	XMLName   xml.Name              `xml:"Types"`
	Defaults  []contentTypeDefault  `xml:"Default"`
	Overrides []contentTypeOverride `xml:"Override"`
}

type contentTypeDefault struct {
	Extension string `xml:"Extension,attr"`
}

type contentTypeOverride struct {
	PartName string `xml:"PartName,attr"`
}

func assertEmbeddedWorkbookContentTypes(t *testing.T, zipPath string) {
	t.Helper()

	reader, err := zip.OpenReader(zipPath)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer reader.Close()

	var contentTypesXML []byte
	for _, file := range reader.File {
		if file.Name != "[Content_Types].xml" {
			continue
		}

		rc, openErr := file.Open()
		if openErr != nil {
			t.Fatalf("open [Content_Types].xml: %v", openErr)
		}
		contentTypesXML, err = io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read [Content_Types].xml: %v", err)
		}
		break
	}

	if len(contentTypesXML) == 0 {
		t.Fatal("[Content_Types].xml not found")
	}

	var parsed contentTypes
	if err := xml.Unmarshal(contentTypesXML, &parsed); err != nil {
		t.Fatalf("unmarshal [Content_Types].xml: %v", err)
	}

	defaults := make(map[string]bool, len(parsed.Defaults))
	for _, def := range parsed.Defaults {
		defaults[strings.ToLower(def.Extension)] = true
	}

	overrides := make(map[string]bool, len(parsed.Overrides))
	for _, override := range parsed.Overrides {
		overrides[strings.TrimPrefix(override.PartName, "/")] = true
	}

	for _, file := range reader.File {
		if !strings.HasPrefix(file.Name, "word/embeddings/") || strings.ToLower(filepath.Ext(file.Name)) != ".xlsx" {
			continue
		}
		if overrides[file.Name] {
			continue
		}

		ext := strings.TrimPrefix(strings.ToLower(filepath.Ext(file.Name)), ".")
		if ext != "" && defaults[ext] {
			continue
		}

		t.Fatalf("missing content type declaration for %s", file.Name)
	}
}

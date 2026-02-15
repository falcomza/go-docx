package docxupdater_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

func TestChartWithCaption(t *testing.T) {
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

	// Insert chart with caption after
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Test Chart",
		Categories: []string{"A", "B", "C"},
		Series: []docxupdater.SeriesData{
			{Name: "Series 1", Values: []float64{10, 20, 30}},
		},
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "Test chart description",
			AutoNumber:  true,
			Position:    docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify caption was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Figure") {
		t.Error("Caption label 'Figure' not found in document.xml")
	}
	if !strings.Contains(docXML, "Test chart description") {
		t.Error("Caption description not found in document.xml")
	}
	// Check for SEQ field (automatic numbering)
	if !strings.Contains(docXML, "SEQ") {
		t.Error("SEQ field for auto-numbering not found in document.xml")
	}
}

func TestChartWithCaptionBefore(t *testing.T) {
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

	// Insert chart with caption before
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Test Chart Before",
		Categories: []string{"X", "Y", "Z"},
		Series: []docxupdater.SeriesData{
			{Name: "Data", Values: []float64{5, 15, 25}},
		},
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionFigure,
			Description: "Caption positioned before chart",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore,
		},
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify caption exists
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Caption positioned before chart") {
		t.Error("Caption description not found in document.xml")
	}
}

func TestTableWithCaption(t *testing.T) {
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

	// Insert table with caption before (typical for tables)
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
			{"Data 3", "Data 4"},
		},
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionTable,
			Description: "Sample data table",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore,
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify caption was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Table") {
		t.Error("Caption label 'Table' not found in document.xml")
	}
	if !strings.Contains(docXML, "Sample data table") {
		t.Error("Caption description not found in document.xml")
	}
	if !strings.Contains(docXML, "SEQ") {
		t.Error("SEQ field for auto-numbering not found in document.xml")
	}
}

func TestTableWithCaptionAfter(t *testing.T) {
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

	// Insert table with caption after (less common but valid)
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Item"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"A", "100"},
			{"B", "200"},
		},
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionTable,
			Description: "Values by item",
			AutoNumber:  true,
			Position:    docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify caption exists
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Values by item") {
		t.Error("Caption description not found in document.xml")
	}
}

func TestCaptionWithManualNumbering(t *testing.T) {
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

	// Insert chart with manual numbering
	err = u.InsertChart(docxupdater.ChartOptions{
		Position:   docxupdater.PositionEnd,
		Title:      "Manual Number Chart",
		Categories: []string{"Q1", "Q2"},
		Series: []docxupdater.SeriesData{
			{Name: "Values", Values: []float64{100, 200}},
		},
		Caption: &docxupdater.CaptionOptions{
			Type:         docxupdater.CaptionFigure,
			Description:  "Manually numbered figure",
			AutoNumber:   false,
			ManualNumber: 42,
			Position:     docxupdater.CaptionAfter,
		},
	})
	if err != nil {
		t.Fatalf("InsertChart failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify manual number
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "42") {
		t.Error("Manual number '42' not found in document.xml")
	}
	if strings.Contains(docXML, "SEQ") {
		t.Error("SEQ field should not be present for manual numbering")
	}
}

func TestCaptionWithCenteredAlignment(t *testing.T) {
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

	// Insert table with centered caption
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Name"},
			{Title: "Score"},
		},
		Rows: [][]string{
			{"Alice", "95"},
			{"Bob", "87"},
		},
		Caption: &docxupdater.CaptionOptions{
			Type:        docxupdater.CaptionTable,
			Description: "Student scores",
			AutoNumber:  true,
			Position:    docxupdater.CaptionBefore,
			Alignment:   docxupdater.CellAlignCenter,
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify centered alignment
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "center") {
		t.Error("Center alignment not found in caption")
	}
}

func TestDefaultCaptionOptions(t *testing.T) {
	// Test default options for Figure
	figureDefaults := docxupdater.DefaultCaptionOptions(docxupdater.CaptionFigure)
	if figureDefaults.Type != docxupdater.CaptionFigure {
		t.Errorf("Expected CaptionFigure, got %v", figureDefaults.Type)
	}
	if figureDefaults.Position != docxupdater.CaptionAfter {
		t.Errorf("Expected CaptionAfter for figures, got %v", figureDefaults.Position)
	}
	if !figureDefaults.AutoNumber {
		t.Error("Expected AutoNumber to be true by default")
	}

	// Test default options for Table
	tableDefaults := docxupdater.DefaultCaptionOptions(docxupdater.CaptionTable)
	if tableDefaults.Type != docxupdater.CaptionTable {
		t.Errorf("Expected CaptionTable, got %v", tableDefaults.Type)
	}
	if tableDefaults.Position != docxupdater.CaptionBefore {
		t.Errorf("Expected CaptionBefore for tables, got %v", tableDefaults.Position)
	}
	if !tableDefaults.AutoNumber {
		t.Error("Expected AutoNumber to be true by default")
	}
}

func TestValidateCaptionOptions(t *testing.T) {
	tests := []struct {
		name    string
		opts    *docxupdater.CaptionOptions
		wantErr bool
	}{
		{
			name: "valid figure caption",
			opts: &docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionFigure,
				Description: "Valid description",
				AutoNumber:  true,
			},
			wantErr: false,
		},
		{
			name: "valid table caption",
			opts: &docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionTable,
				Description: "Valid description",
				AutoNumber:  true,
			},
			wantErr: false,
		},
		{
			name: "invalid caption type",
			opts: &docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionType("Invalid"),
				Description: "Test",
				AutoNumber:  true,
			},
			wantErr: true,
		},
		{
			name: "invalid position",
			opts: &docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionFigure,
				Description: "Test",
				Position:    docxupdater.CaptionPosition("invalid"),
				AutoNumber:  true,
			},
			wantErr: true,
		},
		{
			name: "description too long",
			opts: &docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionFigure,
				Description: strings.Repeat("a", 501),
				AutoNumber:  true,
			},
			wantErr: true,
		},
		{
			name:    "nil caption options",
			opts:    nil,
			wantErr: false, // nil is valid (no caption)
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := docxupdater.ValidateCaptionOptions(tt.opts)
			if (err != nil) != tt.wantErr {
				t.Errorf("ValidateCaptionOptions() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestFormatCaptionText(t *testing.T) {
	tests := []struct {
		name     string
		opts     docxupdater.CaptionOptions
		expected string
	}{
		{
			name: "auto-numbered figure",
			opts: docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionFigure,
				Description: "Test description",
				AutoNumber:  true,
			},
			expected: "Figure #: Test description",
		},
		{
			name: "manually numbered table",
			opts: docxupdater.CaptionOptions{
				Type:         docxupdater.CaptionTable,
				Description:  "Sample table",
				AutoNumber:   false,
				ManualNumber: 5,
			},
			expected: "Table 5: Sample table",
		},
		{
			name: "no description",
			opts: docxupdater.CaptionOptions{
				Type:       docxupdater.CaptionFigure,
				AutoNumber: true,
			},
			expected: "Figure #",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := docxupdater.FormatCaptionText(tt.opts)
			if result != tt.expected {
				t.Errorf("FormatCaptionText() = %q, want %q", result, tt.expected)
			}
		})
	}
}

func TestMultipleChartsCaptionsAutoNumbering(t *testing.T) {
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

	// Insert multiple charts with auto-numbered captions
	for i := 1; i <= 3; i++ {
		err = u.InsertChart(docxupdater.ChartOptions{
			Position:   docxupdater.PositionEnd,
			Title:      "Chart " + string(rune('A'+i-1)),
			Categories: []string{"X", "Y"},
			Series: []docxupdater.SeriesData{
				{Name: "Data", Values: []float64{float64(i * 10), float64(i * 20)}},
			},
			Caption: &docxupdater.CaptionOptions{
				Type:        docxupdater.CaptionFigure,
				Description: "Chart number " + string(rune('A'+i-1)),
				AutoNumber:  true,
				Position:    docxupdater.CaptionAfter,
			},
		})
		if err != nil {
			t.Fatalf("InsertChart %d failed: %v", i, err)
		}
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify all captions exist
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	for i := 0; i < 3; i++ {
		expected := "Chart number " + string(rune('A'+i))
		if !strings.Contains(docXML, expected) {
			t.Errorf("Caption description %q not found in document.xml", expected)
		}
	}
	// Count SEQ fields (should be at least 3)
	seqCount := strings.Count(docXML, "SEQ Figure")
	if seqCount < 3 {
		t.Errorf("Expected at least 3 SEQ fields, found %d", seqCount)
	}
}

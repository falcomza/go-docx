package godocx

import (
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

func TestInsertChartExtended(t *testing.T) {
	tests := []struct {
		name string
		opts ChartOptions
		want error
	}{
		{
			name: "minimal valid chart",
			opts: ChartOptions{
				Categories: []string{"A", "B", "C"},
				Series: []SeriesOptions{
					{Name: "Series1", Values: []float64{1, 2, 3}},
				},
			},
			want: nil,
		},
		{
			name: "empty categories",
			opts: ChartOptions{
				Categories: []string{},
				Series: []SeriesOptions{
					{Name: "Series1", Values: []float64{1, 2, 3}},
				},
			},
			want: fmt.Errorf("dummy"),
		},
		{
			name: "empty series",
			opts: ChartOptions{
				Categories: []string{"A", "B", "C"},
				Series:     []SeriesOptions{},
			},
			want: fmt.Errorf("dummy"),
		},
		{
			name: "mismatched values length",
			opts: ChartOptions{
				Categories: []string{"A", "B", "C"},
				Series: []SeriesOptions{
					{Name: "Series1", Values: []float64{1, 2}}, // Only 2 values
				},
			},
			want: fmt.Errorf("dummy"),
		},
		{
			name: "empty series name",
			opts: ChartOptions{
				Categories: []string{"A", "B", "C"},
				Series: []SeriesOptions{
					{Name: "", Values: []float64{1, 2, 3}},
				},
			},
			want: fmt.Errorf("dummy"),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			templatePath := filepath.Join("templates", "docx_template.docx")
			outputPath := filepath.Join("outputs", "chart_extended_"+tt.name+".docx")

			// Check if template exists
			if _, err := os.Stat(templatePath); os.IsNotExist(err) {
				t.Skip("Template file not found")
			}

			updater, err := New(templatePath)
			if err != nil {
				t.Fatalf("Failed to create updater: %v", err)
			}
			defer updater.Save(outputPath)

			err = updater.InsertChart(tt.opts)

			if tt.want == nil && err != nil {
				t.Errorf("InsertChart() unexpected error = %v", err)
			}
			if tt.want != nil && err == nil {
				t.Errorf("InsertChart() expected error but got none")
			}
		})
	}
}

func TestValidateExtendedChartOptions(t *testing.T) {
	tests := []struct {
		name    string
		opts    ChartOptions
		wantErr bool
	}{
		{
			name: "valid options",
			opts: ChartOptions{
				Categories: []string{"A", "B"},
				Series: []SeriesOptions{
					{Name: "S1", Values: []float64{1, 2}},
				},
			},
			wantErr: false,
		},
		{
			name: "invalid axis min/max",
			opts: ChartOptions{
				Categories: []string{"A", "B"},
				Series: []SeriesOptions{
					{Name: "S1", Values: []float64{1, 2}},
				},
				ValueAxis: &AxisOptions{
					Min: ptrFloat(100),
					Max: ptrFloat(50), // Max < Min
				},
			},
			wantErr: true,
		},
		{
			name: "invalid gap width",
			opts: ChartOptions{
				Categories: []string{"A", "B"},
				Series: []SeriesOptions{
					{Name: "S1", Values: []float64{1, 2}},
				},
				BarChartOptions: &BarChartOptions{
					GapWidth: 600, // > 500
				},
			},
			wantErr: true,
		},
		{
			name: "invalid overlap",
			opts: ChartOptions{
				Categories: []string{"A", "B"},
				Series: []SeriesOptions{
					{Name: "S1", Values: []float64{1, 2}},
				},
				BarChartOptions: &BarChartOptions{
					Overlap: 150, // > 100
				},
			},
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validateChartOptions(tt.opts)
			if (err != nil) != tt.wantErr {
				t.Errorf("validateChartOptions() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestApplyExtendedChartDefaults(t *testing.T) {
	t.Run("applies chart kind default", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S1", Values: []float64{1}}},
		}
		result := applyChartDefaults(opts)
		if result.ChartKind != ChartKindColumn {
			t.Errorf("Expected ChartKindColumn, got %v", result.ChartKind)
		}
	})

	t.Run("applies dimensions default", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S1", Values: []float64{1}}},
		}
		result := applyChartDefaults(opts)
		if result.Width == 0 || result.Height == 0 {
			t.Errorf("Expected non-zero dimensions, got Width=%d, Height=%d", result.Width, result.Height)
		}
	})

	t.Run("applies legend defaults", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S1", Values: []float64{1}}},
		}
		result := applyChartDefaults(opts)
		if result.Legend == nil {
			t.Fatal("Expected legend to be initialized")
		}
		if result.Legend.Position != "r" {
			t.Errorf("Expected legend position 'r', got %v", result.Legend.Position)
		}
	})

	t.Run("applies axis defaults", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S1", Values: []float64{1}}},
		}
		result := applyChartDefaults(opts)

		if result.CategoryAxis == nil {
			t.Fatal("Expected category axis to be initialized")
		}
		if result.CategoryAxis.Position != AxisPositionBottom {
			t.Errorf("Expected category axis position bottom, got %v", result.CategoryAxis.Position)
		}

		if result.ValueAxis == nil {
			t.Fatal("Expected value axis to be initialized")
		}
		if result.ValueAxis.Position != AxisPositionLeft {
			t.Errorf("Expected value axis position left, got %v", result.ValueAxis.Position)
		}
		if !result.ValueAxis.MajorGridlines {
			t.Error("Expected value axis major gridlines to be true")
		}
	})

	t.Run("applies bar chart defaults", func(t *testing.T) {
		opts := ChartOptions{
			ChartKind:  ChartKindColumn,
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S1", Values: []float64{1}}},
		}
		result := applyChartDefaults(opts)

		if result.BarChartOptions == nil {
			t.Fatal("Expected bar chart options to be initialized")
		}
		if result.BarChartOptions.Direction != BarDirectionColumn {
			t.Errorf("Expected column direction, got %v", result.BarChartOptions.Direction)
		}
		if result.BarChartOptions.Grouping != BarGroupingClustered {
			t.Errorf("Expected clustered grouping, got %v", result.BarChartOptions.Grouping)
		}
		if result.BarChartOptions.GapWidth != 150 {
			t.Errorf("Expected gap width 150, got %d", result.BarChartOptions.GapWidth)
		}
	})

	t.Run("applies chart properties defaults", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S1", Values: []float64{1}}},
		}
		result := applyChartDefaults(opts)

		if result.Properties == nil {
			t.Fatal("Expected properties to be initialized")
		}
		if result.Properties.Style != ChartStyle2 {
			t.Errorf("Expected ChartStyle2, got %v", result.Properties.Style)
		}
		if result.Properties.Language != "en-US" {
			t.Errorf("Expected language en-US, got %v", result.Properties.Language)
		}
		if result.Properties.DisplayBlanksAs != "gap" {
			t.Errorf("Expected display blanks as gap, got %v", result.Properties.DisplayBlanksAs)
		}
	})
}

func TestGenerateExtendedChartXML(t *testing.T) {
	t.Run("generates valid XML with minimal options", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A", "B", "C"},
			Series: []SeriesOptions{
				{Name: "Series1", Values: []float64{1, 2, 3}},
			},
		}
		opts = applyChartDefaults(opts)

		xml := generateChartXML(opts)
		xmlStr := string(xml)

		// Verify XML declaration
		if !containsString(xmlStr, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>") {
			t.Error("Missing XML declaration")
		}

		// Verify namespaces
		requiredNamespaces := []string{
			"xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"",
			"xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
			"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
		}
		for _, ns := range requiredNamespaces {
			if !containsString(xmlStr, ns) {
				t.Errorf("Missing namespace: %s", ns)
			}
		}

		// Verify chart properties
		if !containsString(xmlStr, "<c:date1904") {
			t.Error("Missing date1904 property")
		}
		if !containsString(xmlStr, "<c:lang") {
			t.Error("Missing lang property")
		}
		if !containsString(xmlStr, "<c:roundedCorners") {
			t.Error("Missing roundedCorners property")
		}
	})

	t.Run("generates XML with custom axes", func(t *testing.T) {
		minVal := 0.0
		maxVal := 100.0

		opts := ChartOptions{
			Categories: []string{"A", "B"},
			Series: []SeriesOptions{
				{Name: "Series1", Values: []float64{50, 75}},
			},
			ValueAxis: &AxisOptions{
				Title:        "Custom Axis",
				Min:          &minVal,
				Max:          &maxVal,
				NumberFormat: "#,##0",
			},
		}
		opts = applyChartDefaults(opts)

		xml := string(generateChartXML(opts))

		if !containsString(xml, "Custom Axis") {
			t.Error("Missing custom axis title")
		}
		if !containsString(xml, "<c:min val=\"0\"") {
			t.Error("Missing min value")
		}
		if !containsString(xml, "<c:max val=\"100\"") {
			t.Error("Missing max value")
		}
	})

	t.Run("generates XML with data labels", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A", "B"},
			Series: []SeriesOptions{
				{Name: "Series1", Values: []float64{10, 20}},
			},
			DataLabels: &DataLabelOptions{
				ShowValue:   true,
				ShowPercent: false,
				Position:    DataLabelOutsideEnd,
			},
		}
		opts = applyChartDefaults(opts)

		xml := string(generateChartXML(opts))

		if !containsString(xml, "<c:showVal val=\"1\"") {
			t.Error("Missing showVal=1 for data labels")
		}
		if !containsString(xml, "<c:dLblPos val=\"outEnd\"") {
			t.Error("Missing data label position")
		}
	})

	t.Run("generates XML with custom colors", func(t *testing.T) {
		opts := ChartOptions{
			Categories: []string{"A", "B"},
			Series: []SeriesOptions{
				{
					Name:   "Series1",
					Values: []float64{10, 20},
					Color:  "FF0000", // Red
				},
			},
		}
		opts = applyChartDefaults(opts)

		xml := string(generateChartXML(opts))

		if !containsString(xml, "FF0000") {
			t.Error("Missing custom color")
		}
	})
}

func TestChartTypeXMLGeneration(t *testing.T) {
	baseOpts := ChartOptions{
		Categories: []string{"A", "B"},
		Series: []SeriesOptions{
			{Name: "Series1", Values: []float64{10, 20}},
		},
	}

	tests := []struct {
		name      string
		chartKind ChartKind
		contains  string
	}{
		{"column chart", ChartKindColumn, "<c:barChart>"},
		{"bar chart", ChartKindBar, "<c:barChart>"},
		{"line chart", ChartKindLine, "<c:lineChart>"},
		{"pie chart", ChartKindPie, "<c:pieChart>"},
		{"area chart", ChartKindArea, "<c:areaChart>"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			opts := baseOpts
			opts.ChartKind = tt.chartKind
			opts = applyChartDefaults(opts)

			xml := string(generateChartXML(opts))

			if !containsString(xml, tt.contains) {
				t.Errorf("Expected XML to contain %s", tt.contains)
			}
		})
	}
}

func TestLineChartSpecificOptions(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindLine,
		Categories: []string{"A", "B", "C"},
		Series: []SeriesOptions{
			{
				Name:        "Series1",
				Values:      []float64{10, 20, 15},
				Smooth:      true,
				ShowMarkers: true,
			},
		},
	}
	opts = applyChartDefaults(opts)

	xml := string(generateChartXML(opts))

	if !containsString(xml, "<c:smooth val=\"1\"") {
		t.Error("Missing smooth line option")
	}
	if !containsString(xml, "<c:symbol val=\"circle\"") {
		t.Error("Missing marker symbol")
	}
}

func TestBoolToInt(t *testing.T) {
	tests := []struct {
		input bool
		want  int
	}{
		{true, 1},
		{false, 0},
	}

	for _, tt := range tests {
		t.Run("", func(t *testing.T) {
			got := boolToInt(tt.input)
			if got != tt.want {
				t.Errorf("boolToInt(%v) = %d, want %d", tt.input, got, tt.want)
			}
		})
	}
}

func TestGetChartCountNonChartFiles(t *testing.T) {
	// Create a docx whose word/charts/ dir only has non-chart files
	dir := t.TempDir()
	chartsDir := filepath.Join(dir, "word", "charts")
	if err := os.MkdirAll(chartsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	// Write a non-chart file to the directory
	if err := os.WriteFile(filepath.Join(chartsDir, "colors1.xml"), []byte("<colors/>"), 0o644); err != nil {
		t.Fatal(err)
	}

	u := &Updater{tempDir: dir}
	count, err := u.GetChartCount()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("expected 0, got %d", count)
	}
}

// Helper functions

func ptrFloat(f float64) *float64 {
	return &f
}

func containsString(s, substr string) bool {
	return len(s) > 0 && len(substr) > 0 &&
		(s == substr || len(s) >= len(substr) && findSubstring(s, substr))
}

func findSubstring(s, substr string) bool {
	if len(substr) > len(s) {
		return false
	}
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}

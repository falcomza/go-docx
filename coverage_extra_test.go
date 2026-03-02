package godocx

import (
	"strings"
	"testing"
)

// Tests for validateAxisOptions (22% coverage -> should be near 100%)
func TestValidateAxisOptions(t *testing.T) {
	minVal := 10.0
	maxVal := 100.0
	majorUnit := 20.0
	minorUnit := 5.0
	badMinorUnit := 25.0
	negUnit := -1.0

	tests := []struct {
		name    string
		axis    *AxisOptions
		wantErr bool
	}{
		{
			name: "valid axis with all options",
			axis: &AxisOptions{
				Min:       &minVal,
				Max:       &maxVal,
				MajorUnit: &majorUnit,
				MinorUnit: &minorUnit,
			},
			wantErr: false,
		},
		{
			name: "min >= max",
			axis: &AxisOptions{
				Min: &maxVal,
				Max: &minVal,
			},
			wantErr: true,
		},
		{
			name: "min == max",
			axis: &AxisOptions{
				Min: &minVal,
				Max: &minVal,
			},
			wantErr: true,
		},
		{
			name: "negative major unit",
			axis: &AxisOptions{
				MajorUnit: &negUnit,
			},
			wantErr: true,
		},
		{
			name: "negative minor unit",
			axis: &AxisOptions{
				MinorUnit: &negUnit,
			},
			wantErr: true,
		},
		{
			name: "minor >= major",
			axis: &AxisOptions{
				MajorUnit: &majorUnit,
				MinorUnit: &badMinorUnit,
			},
			wantErr: true,
		},
		{
			name:    "empty axis is valid",
			axis:    &AxisOptions{},
			wantErr: false,
		},
		{
			name: "only min is valid",
			axis: &AxisOptions{
				Min: &minVal,
			},
			wantErr: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validateAxisOptions("TestAxis", tt.axis)
			if tt.wantErr && err == nil {
				t.Error("expected error")
			}
			if !tt.wantErr && err != nil {
				t.Errorf("unexpected error: %v", err)
			}
		})
	}
}

// Tests for columnLetter (already has some coverage but let's be thorough)
func TestColumnLetter(t *testing.T) {
	tests := []struct {
		input int
		want  string
	}{
		{1, "A"},
		{2, "B"},
		{3, "C"},
		{26, "Z"},
		{27, "AA"},
		{28, "AB"},
		{52, "AZ"},
		{53, "BA"},
	}

	for _, tt := range tests {
		got := columnLetter(tt.input)
		if got != tt.want {
			t.Errorf("columnLetter(%d) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestGenerateBarChartXML(t *testing.T) {
	opts := ChartOptions{
		ChartKind: ChartKindColumn,
		Categories: []string{"Q1", "Q2"},
		Series: []SeriesOptions{
			{Name: "Revenue", Values: []float64{100, 200}},
		},
	}
	opts = applyChartDefaults(opts)

	result := generateBarChartXML(opts)

	if !strings.Contains(result, "<c:barChart>") {
		t.Error("expected barChart element")
	}
	if !strings.Contains(result, `<c:barDir val="col"/>`) {
		t.Error("expected column direction")
	}
	if !strings.Contains(result, "Revenue") {
		t.Error("expected series name")
	}
}

func TestGenerateLineChartXML(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindLine,
		Categories: []string{"Jan", "Feb"},
		Series: []SeriesOptions{
			{Name: "Sales", Values: []float64{50, 75}, Smooth: true, ShowMarkers: true},
		},
	}
	opts = applyChartDefaults(opts)

	result := generateLineChartXML(opts)

	if !strings.Contains(result, "<c:lineChart>") {
		t.Error("expected lineChart element")
	}
	if !strings.Contains(result, `<c:smooth val="1"/>`) {
		t.Error("expected smooth line")
	}
	if !strings.Contains(result, `<c:symbol val="circle"/>`) {
		t.Error("expected circle markers")
	}
}

func TestGeneratePieChartXML(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindPie,
		Categories: []string{"A", "B", "C"},
		Series: []SeriesOptions{
			{Name: "Share", Values: []float64{30, 50, 20}},
		},
	}
	opts = applyChartDefaults(opts)

	result := generatePieChartXML(opts)

	if !strings.Contains(result, "<c:pieChart>") {
		t.Error("expected pieChart element")
	}
	if !strings.Contains(result, `<c:varyColors val="1"/>`) {
		t.Error("expected varyColors for pie")
	}
}

func TestGenerateAreaChartXML(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindArea,
		Categories: []string{"X", "Y"},
		Series: []SeriesOptions{
			{Name: "Area1", Values: []float64{10, 20}},
		},
	}
	opts = applyChartDefaults(opts)

	result := generateAreaChartXML(opts)

	if !strings.Contains(result, "<c:areaChart>") {
		t.Error("expected areaChart element")
	}
}

func TestGenerateSeriesXML_WithColor(t *testing.T) {
	series := SeriesOptions{
		Name:              "Colored",
		Values:            []float64{10, 20},
		Color:             "#FF0000",
		InvertIfNegative:  true,
	}
	opts := ChartOptions{
		ChartKind:  ChartKindColumn,
		Categories: []string{"A", "B"},
		Series:     []SeriesOptions{series},
	}
	opts = applyChartDefaults(opts)

	result := generateSeriesXML(0, series, opts)

	if !strings.Contains(result, "FF0000") {
		t.Error("expected color in output")
	}
	if !strings.Contains(result, `<c:invertIfNegative val="1"/>`) {
		t.Error("expected invertIfNegative")
	}
}

func TestGenerateDataLabelsXML(t *testing.T) {
	labels := &DataLabelOptions{
		ShowValue:        true,
		ShowCategoryName: true,
		ShowSeriesName:   true,
		ShowPercent:      true,
		ShowLegendKey:    true,
		ShowLeaderLines:  true,
		Position:         "outEnd",
	}

	result := generateDataLabelsXML(labels)

	if !strings.Contains(result, `<c:showVal val="1"/>`) {
		t.Error("expected showVal")
	}
	if !strings.Contains(result, `<c:showCatName val="1"/>`) {
		t.Error("expected showCatName")
	}
	if !strings.Contains(result, `<c:showSerName val="1"/>`) {
		t.Error("expected showSerName")
	}
	if !strings.Contains(result, `<c:showPercent val="1"/>`) {
		t.Error("expected showPercent")
	}
	if !strings.Contains(result, `<c:showLegendKey val="1"/>`) {
		t.Error("expected showLegendKey")
	}
	if !strings.Contains(result, `<c:showLeaderLines val="1"/>`) {
		t.Error("expected showLeaderLines")
	}
	if !strings.Contains(result, `<c:dLblPos val="outEnd"/>`) {
		t.Error("expected dLblPos")
	}
}

func TestGenerateCategoryAxisXML_Extended(t *testing.T) {
	minVal := 0.0
	maxVal := 100.0
	crossAt := 50.0

	axis := &AxisOptions{
		Title:          "X Axis",
		TitleOverlay:   true,
		Visible:        true,
		Position:       AxisPositionBottom,
		MajorGridlines: true,
		MinorGridlines: true,
		Min:            &minVal,
		Max:            &maxVal,
		CrossesAt:      &crossAt,
		NumberFormat:   "0.00",
		MajorTickMark:  TickMarkOut,
		MinorTickMark:  TickMarkIn,
		TickLabelPos:   TickLabelNextTo,
	}

	result := generateCategoryAxisXML(axis)

	if !strings.Contains(result, "X Axis") {
		t.Error("expected axis title")
	}
	if !strings.Contains(result, `<c:min val="0"/>`) {
		t.Error("expected min value")
	}
	if !strings.Contains(result, `<c:max val="100"/>`) {
		t.Error("expected max value")
	}
	if !strings.Contains(result, `<c:crossesAt val="50"/>`) {
		t.Error("expected crossesAt")
	}
	if !strings.Contains(result, `<c:majorGridlines/>`) {
		t.Error("expected major gridlines")
	}
	if !strings.Contains(result, `<c:minorGridlines/>`) {
		t.Error("expected minor gridlines")
	}
}

func TestGenerateValueAxisXML_Extended(t *testing.T) {
	majorUnit := 10.0
	minorUnit := 2.0

	axis := &AxisOptions{
		Title:          "Y Axis",
		Visible:        true,
		Position:       AxisPositionLeft,
		MajorGridlines: true,
		MinorGridlines: true,
		MajorUnit:      &majorUnit,
		MinorUnit:      &minorUnit,
		NumberFormat:   "#,##0",
		MajorTickMark:  TickMarkCross,
		MinorTickMark:  TickMarkNone,
		TickLabelPos:   TickLabelNextTo,
	}

	result := generateValueAxisXML(axis)

	if !strings.Contains(result, "Y Axis") {
		t.Error("expected axis title")
	}
	if !strings.Contains(result, `<c:majorUnit val="10"/>`) {
		t.Error("expected major unit")
	}
	if !strings.Contains(result, `<c:minorUnit val="2"/>`) {
		t.Error("expected minor unit")
	}
	if !strings.Contains(result, `<c:majorGridlines/>`) {
		t.Error("expected major gridlines")
	}
}

func TestValidateChartOptions_EdgeCases(t *testing.T) {
	t.Run("empty categories", func(t *testing.T) {
		err := validateChartOptions(ChartOptions{
			Series: []SeriesOptions{{Name: "S", Values: []float64{1}}},
		})
		if err == nil {
			t.Error("expected error for empty categories")
		}
	})

	t.Run("empty series", func(t *testing.T) {
		err := validateChartOptions(ChartOptions{
			Categories: []string{"A"},
		})
		if err == nil {
			t.Error("expected error for empty series")
		}
	})

	t.Run("empty series name", func(t *testing.T) {
		err := validateChartOptions(ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "  ", Values: []float64{1}}},
		})
		if err == nil {
			t.Error("expected error for empty series name")
		}
	})

	t.Run("values length mismatch", func(t *testing.T) {
		err := validateChartOptions(ChartOptions{
			Categories: []string{"A", "B"},
			Series:     []SeriesOptions{{Name: "S", Values: []float64{1}}},
		})
		if err == nil {
			t.Error("expected error for values length mismatch")
		}
	})

	t.Run("bar chart options validation - gap width", func(t *testing.T) {
		err := validateChartOptions(ChartOptions{
			Categories:      []string{"A"},
			Series:          []SeriesOptions{{Name: "S", Values: []float64{1}}},
			BarChartOptions: &BarChartOptions{GapWidth: 600},
		})
		if err == nil {
			t.Error("expected error for gap width > 500")
		}
	})

	t.Run("bar chart options validation - overlap", func(t *testing.T) {
		err := validateChartOptions(ChartOptions{
			Categories:      []string{"A"},
			Series:          []SeriesOptions{{Name: "S", Values: []float64{1}}},
			BarChartOptions: &BarChartOptions{Overlap: 150},
		})
		if err == nil {
			t.Error("expected error for overlap > 100")
		}
	})
}

func TestApplyChartDefaults(t *testing.T) {
	t.Run("bar chart defaults", func(t *testing.T) {
		// ChartKindBar is the horizontal-bar sentinel ("bar"); applyChartDefaults
		// should therefore set the direction to BarDirectionBar ("bar").
		opts := applyChartDefaults(ChartOptions{
			ChartKind:  ChartKindBar,
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S", Values: []float64{1}}},
		})
		if opts.BarChartOptions.Direction != BarDirectionBar {
			t.Errorf("expected bar direction for ChartKindBar, got %s", opts.BarChartOptions.Direction)
		}
	})

	t.Run("column chart defaults", func(t *testing.T) {
		opts := applyChartDefaults(ChartOptions{
			ChartKind:  ChartKindColumn,
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S", Values: []float64{1}}},
		})
		if opts.BarChartOptions.Direction != BarDirectionColumn {
			t.Errorf("expected column direction, got %s", opts.BarChartOptions.Direction)
		}
	})

	t.Run("backward compat legend", func(t *testing.T) {
		opts := applyChartDefaults(ChartOptions{
			ShowLegend:     true,
			LegendPosition: "b",
			Categories:     []string{"A"},
			Series:         []SeriesOptions{{Name: "S", Values: []float64{1}}},
		})
		if !opts.Legend.Show {
			t.Error("expected legend to be shown")
		}
		if opts.Legend.Position != "b" {
			t.Errorf("expected position 'b', got %s", opts.Legend.Position)
		}
	})

	t.Run("backward compat axis titles", func(t *testing.T) {
		opts := applyChartDefaults(ChartOptions{
			CategoryAxisTitle: "X Label",
			ValueAxisTitle:    "Y Label",
			Categories:        []string{"A"},
			Series:            []SeriesOptions{{Name: "S", Values: []float64{1}}},
		})
		if opts.CategoryAxis.Title != "X Label" {
			t.Errorf("expected X Label, got %s", opts.CategoryAxis.Title)
		}
		if opts.ValueAxis.Title != "Y Label" {
			t.Errorf("expected Y Label, got %s", opts.ValueAxis.Title)
		}
	})

	t.Run("data labels defaults", func(t *testing.T) {
		opts := applyChartDefaults(ChartOptions{
			Categories: []string{"A"},
			Series:     []SeriesOptions{{Name: "S", Values: []float64{1}}},
			DataLabels: &DataLabelOptions{ShowValue: true},
		})
		if opts.DataLabels.Position != DataLabelBestFit {
			t.Errorf("expected default position bestFit, got %s", opts.DataLabels.Position)
		}
	})
}

func TestGenerateSheetXML_Scatter(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindScatter,
		Categories: []string{"P1", "P2"},
		Series: []SeriesOptions{
			{
				Name:    "Data",
				Values:  []float64{10, 20},
				XValues: []float64{1.5, 3.0},
			},
		},
	}

	result := string(generateSheetXML(opts))

	if !strings.Contains(result, "X Values") {
		t.Error("expected X Values header")
	}
	if !strings.Contains(result, "1.5") {
		t.Error("expected X value 1.5")
	}
	if !strings.Contains(result, "3") {
		t.Error("expected X value 3.0")
	}
}

func TestGenerateSheetXML_NoXValues(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindColumn,
		Categories: []string{"Q1", "Q2"},
		Series: []SeriesOptions{
			{Name: "Rev", Values: []float64{100, 200}},
		},
	}

	result := string(generateSheetXML(opts))

	if strings.Contains(result, "X Values") {
		t.Error("should not have X Values header for non-scatter")
	}
	if !strings.Contains(result, "Q1") {
		t.Error("expected category Q1")
	}
}

func TestGenerateTitleXML(t *testing.T) {
	result := generateTitleXML("My Chart", false)
	if !strings.Contains(result, "My Chart") {
		t.Error("expected title text")
	}
	if !strings.Contains(result, `<c:overlay val="0"/>`) {
		t.Error("expected no overlay")
	}

	resultOverlay := generateTitleXML("Overlay", true)
	if !strings.Contains(resultOverlay, `<c:overlay val="1"/>`) {
		t.Error("expected overlay")
	}
}

func TestGenerateLegendXML(t *testing.T) {
	legend := &LegendOptions{
		Show:     true,
		Position: "b",
		Overlay:  true,
	}

	result := generateLegendXML(legend)

	if !strings.Contains(result, `<c:legendPos val="b"/>`) {
		t.Error("expected legend position bottom")
	}
	if !strings.Contains(result, `<c:overlay val="1"/>`) {
		t.Error("expected overlay")
	}
}

func TestFindAllSeriesTags(t *testing.T) {
	content := `<c:barChart><c:ser><c:idx val="0"/></c:ser><c:ser><c:idx val="1"/></c:ser><c:axId val="1"/></c:barChart>`

	positions := findAllSeriesTags(content, "c:")
	if len(positions) != 2 {
		t.Errorf("expected 2 series tags, got %d", len(positions))
	}
}

func TestCopyNonSeriesElements(t *testing.T) {
	content := `<c:barChart><c:ser><c:idx/></c:ser><c:axId val="123"/><c:gapWidth val="150"/></c:barChart>`

	result := copyNonSeriesElements(content, "c:")

	if !strings.Contains(result, "axId") {
		t.Error("expected axId in result")
	}
	if !strings.Contains(result, "gapWidth") {
		t.Error("expected gapWidth in result")
	}
}

func TestApplyAxisDefaults(t *testing.T) {
	t.Run("category axis defaults", func(t *testing.T) {
		axis := applyAxisDefaults(nil, true)
		if axis.Position != AxisPositionBottom {
			t.Errorf("expected bottom position, got %s", axis.Position)
		}
		if axis.MajorGridlines {
			t.Error("category axis should not have major gridlines by default")
		}
	})

	t.Run("value axis defaults", func(t *testing.T) {
		axis := applyAxisDefaults(nil, false)
		if axis.Position != AxisPositionLeft {
			t.Errorf("expected left position, got %s", axis.Position)
		}
		if !axis.MajorGridlines {
			t.Error("value axis should have major gridlines by default")
		}
	})

	t.Run("preserves existing values", func(t *testing.T) {
		axis := applyAxisDefaults(&AxisOptions{
			Position: "r",
		}, false)
		if axis.Position != "r" {
			t.Error("should preserve existing position")
		}
	})
}

package godocx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// ChartKind defines the type of chart
type ChartKind string

const (
	// ChartKindColumn creates a column chart (vertical bars, <c:barDir val="col"/>).
	ChartKindColumn ChartKind = "column"
	// ChartKindBar creates a bar chart (horizontal bars, <c:barDir val="bar"/>).
	// Although both column and bar charts emit a <c:barChart> element in OpenXML,
	// they are kept as distinct constants so callers do not need to set
	// BarChartOptions.Direction manually.
	ChartKindBar     ChartKind = "bar"
	ChartKindLine    ChartKind = "lineChart"    // Line chart
	ChartKindPie     ChartKind = "pieChart"     // Pie chart
	ChartKindArea    ChartKind = "areaChart"    // Area chart
	ChartKindScatter ChartKind = "scatterChart" // Scatter chart (XY chart)
)

// ChartOptions defines comprehensive options for chart creation
type ChartOptions struct {
	// Position where to insert the chart
	Position InsertPosition
	Anchor   string // Text anchor for relative positioning

	// Chart type (default: Column)
	ChartKind ChartKind

	// Chart titles
	Title             string // Main chart title
	TitleOverlay      bool   // Overlays the title on the chart area
	CategoryAxisTitle string // X-axis title (horizontal axis) — backward compat, prefer CategoryAxis.Title
	ValueAxisTitle    string // Y-axis title (vertical axis) — backward compat, prefer ValueAxis.Title

	// Data
	Categories []string        // Category labels (X-axis)
	Series     []SeriesOptions // Data series with names and values
	// Deprecated: Use Legend.Show instead.
	ShowLegend bool // Show legend — backward compat, prefer Legend.Show
	// Deprecated: Use Legend.Position instead.
	LegendPosition string // Legend position — backward compat, prefer Legend.Position

	// Chart dimensions (default: spans between margins)
	Width  int // Width in EMUs (English Metric Units), 0 for default (6099523 = ~6.5")
	Height int // Height in EMUs, 0 for default (3340467 = ~3.5")

	// Caption options (nil for no caption)
	Caption *CaptionOptions

	// Extended axis customization (nil = auto defaults)
	CategoryAxis *AxisOptions
	ValueAxis    *AxisOptions

	// Legend customization (nil = derives from ShowLegend/LegendPosition)
	Legend *LegendOptions

	// Default data labels for all series (nil = no labels)
	DataLabels *DataLabelOptions

	// Chart-level rendering properties (nil = library defaults)
	Properties *ChartProperties

	// Bar/column-specific options (nil = clustered column defaults)
	BarChartOptions *BarChartOptions

	// Scatter chart-specific options (nil = marker defaults)
	ScatterChartOptions *ScatterChartOptions
}

// InsertChart creates a new chart and inserts it into the document
func (u *Updater) InsertChart(opts ChartOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Validate options
	if err := validateChartOptions(opts); err != nil {
		return fmt.Errorf("invalid chart options: %w", err)
	}

	// Apply defaults
	opts = applyChartDefaults(opts)

	// Find next available chart index
	chartIndex := u.findNextChartIndex()

	// Create chart XML file
	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	if err := u.createChartXML(chartPath, opts); err != nil {
		return fmt.Errorf("create chart xml: %w", err)
	}

	// Create embedded workbook
	workbookPath := filepath.Join(u.tempDir, "word", "embeddings", fmt.Sprintf("Microsoft_Excel_Worksheet%d.xlsx", chartIndex))
	if err := u.createEmbeddedWorkbook(workbookPath, opts); err != nil {
		return fmt.Errorf("create embedded workbook: %w", err)
	}

	// Create chart relationships file
	chartRelsPath := filepath.Join(u.tempDir, "word", "charts", "_rels", fmt.Sprintf("chart%d.xml.rels", chartIndex))
	if err := u.createChartRelationships(chartRelsPath, workbookPath); err != nil {
		return fmt.Errorf("create chart relationships: %w", err)
	}

	// Add chart relationship to document.xml.rels
	relID, err := u.addChartRelationship(chartIndex)
	if err != nil {
		return fmt.Errorf("add chart relationship: %w", err)
	}

	// Insert chart drawing into document
	if err := u.insertChartDrawing(chartIndex, relID, opts); err != nil {
		return fmt.Errorf("insert chart drawing: %w", err)
	}

	// Update content types
	if err := u.addContentTypeOverride(chartIndex); err != nil {
		return fmt.Errorf("add content type: %w", err)
	}
	if err := u.addImageContentType(".xlsx", XLSXContentType); err != nil {
		return fmt.Errorf("add workbook content type: %w", err)
	}

	return nil
}

// validateChartOptions validates chart creation options
func validateChartOptions(opts ChartOptions) error {
	if len(opts.Categories) == 0 {
		return fmt.Errorf("categories cannot be empty")
	}
	if len(opts.Series) == 0 {
		return fmt.Errorf("at least one series is required")
	}

	// Validate series
	for i, series := range opts.Series {
		if strings.TrimSpace(series.Name) == "" {
			return fmt.Errorf("series[%d] name cannot be empty", i)
		}
		if len(series.Values) != len(opts.Categories) {
			return fmt.Errorf("series[%d] values length (%d) must match categories length (%d)", i, len(series.Values), len(opts.Categories))
		}
	}

	// Validate axes if provided
	if opts.CategoryAxis != nil {
		if err := validateAxisOptions("CategoryAxis", opts.CategoryAxis); err != nil {
			return err
		}
	}
	if opts.ValueAxis != nil {
		if err := validateAxisOptions("ValueAxis", opts.ValueAxis); err != nil {
			return err
		}
	}

	// Validate bar chart options if provided
	if opts.BarChartOptions != nil {
		if opts.BarChartOptions.GapWidth < 0 || opts.BarChartOptions.GapWidth > 500 {
			return fmt.Errorf("BarChartOptions.GapWidth must be between 0 and 500")
		}
		if opts.BarChartOptions.Overlap < -100 || opts.BarChartOptions.Overlap > 100 {
			return fmt.Errorf("BarChartOptions.Overlap must be between -100 and 100")
		}
	}

	return nil
}

// validateAxisOptions validates axis options
func validateAxisOptions(name string, axis *AxisOptions) error {
	if axis.Min != nil && axis.Max != nil && *axis.Min >= *axis.Max {
		return fmt.Errorf("%s: Min must be less than Max", name)
	}
	if axis.MajorUnit != nil && *axis.MajorUnit <= 0 {
		return fmt.Errorf("%s: MajorUnit must be positive", name)
	}
	if axis.MinorUnit != nil && *axis.MinorUnit <= 0 {
		return fmt.Errorf("%s: MinorUnit must be positive", name)
	}
	if axis.MajorUnit != nil && axis.MinorUnit != nil && *axis.MinorUnit >= *axis.MajorUnit {
		return fmt.Errorf("%s: MinorUnit must be less than MajorUnit", name)
	}
	return nil
}

// applyChartDefaults sets default values for unspecified options
func applyChartDefaults(opts ChartOptions) ChartOptions {
	// Basic defaults
	if opts.ChartKind == "" {
		opts.ChartKind = ChartKindColumn
	}
	if opts.Width == 0 {
		opts.Width = 6099523 // ~6.5 inches
	}
	if opts.Height == 0 {
		opts.Height = 3340467 // ~3.5 inches
	}

	// Apply legend defaults — honour backward-compat ShowLegend/LegendPosition fields
	if opts.Legend == nil {
		pos := opts.LegendPosition
		if pos == "" {
			pos = "r"
		}
		opts.Legend = &LegendOptions{
			Show:     opts.ShowLegend,
			Position: pos,
			Overlay:  false,
		}
	}

	// Apply category axis defaults — honour backward-compat CategoryAxisTitle
	if opts.CategoryAxis == nil {
		opts.CategoryAxis = &AxisOptions{}
		if opts.CategoryAxisTitle != "" {
			opts.CategoryAxis.Title = opts.CategoryAxisTitle
		}
	}
	opts.CategoryAxis = applyAxisDefaults(opts.CategoryAxis, true)

	// Apply value axis defaults — honour backward-compat ValueAxisTitle
	if opts.ValueAxis == nil {
		opts.ValueAxis = &AxisOptions{}
		if opts.ValueAxisTitle != "" {
			opts.ValueAxis.Title = opts.ValueAxisTitle
		}
	}
	opts.ValueAxis = applyAxisDefaults(opts.ValueAxis, false)

	// Apply chart properties defaults
	if opts.Properties == nil {
		opts.Properties = &ChartProperties{}
	}
	if opts.Properties.Style == 0 {
		opts.Properties.Style = ChartStyle2
	}
	if opts.Properties.Language == "" {
		opts.Properties.Language = "en-US"
	}
	if opts.Properties.DisplayBlanksAs == "" {
		opts.Properties.DisplayBlanksAs = "gap"
	}
	opts.Properties.PlotVisibleOnly = true // Always true

	// Apply bar chart defaults if chart is bar/column type
	if opts.ChartKind == ChartKindColumn || opts.ChartKind == ChartKindBar {
		if opts.BarChartOptions == nil {
			opts.BarChartOptions = &BarChartOptions{}
		}
		if opts.BarChartOptions.Direction == "" {
			if opts.ChartKind == ChartKindColumn {
				opts.BarChartOptions.Direction = BarDirectionColumn
			} else {
				opts.BarChartOptions.Direction = BarDirectionBar
			}
		}
		if opts.BarChartOptions.Grouping == "" {
			opts.BarChartOptions.Grouping = BarGroupingClustered
		}
		if opts.BarChartOptions.GapWidth == 0 {
			opts.BarChartOptions.GapWidth = 150
		}
	}

	// Apply data label defaults if specified
	if opts.DataLabels != nil {
		if opts.DataLabels.Position == "" {
			opts.DataLabels.Position = DataLabelBestFit
		}
	}

	return opts
}

// createChartXML generates the chart XML file
func (u *Updater) createChartXML(chartPath string, opts ChartOptions) error {
	if err := os.MkdirAll(filepath.Dir(chartPath), 0o755); err != nil {
		return fmt.Errorf("create charts directory: %w", err)
	}

	xml := generateChartXML(opts)

	if err := atomicWriteFile(chartPath, xml, 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

// generateChartXML creates the chart XML content with all extended options
func generateChartXML(opts ChartOptions) []byte {
	var buf bytes.Buffer

	// XML declaration with newline (Word requires this)
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")

	// chartSpace with all standard namespaces
	buf.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`)
	buf.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">`)

	// Chart properties
	buf.WriteString(fmt.Sprintf(`<c:date1904 val="%d"/>`, boolToInt(opts.Properties.Date1904)))
	buf.WriteString(fmt.Sprintf(`<c:lang val="%s"/>`, opts.Properties.Language))
	buf.WriteString(fmt.Sprintf(`<c:roundedCorners val="%d"/>`, boolToInt(opts.Properties.RoundedCorners)))

	// Chart style (if not default)
	if opts.Properties.Style > 0 {
		buf.WriteString(fmt.Sprintf(`<c:style val="%d"/>`, opts.Properties.Style))
	}

	buf.WriteString(`<c:chart>`)

	// Chart title
	if opts.Title != "" {
		buf.WriteString(generateTitleXML(opts.Title, opts.TitleOverlay))
	}

	buf.WriteString(`<c:autoTitleDeleted val="0"/>`)
	buf.WriteString(`<c:plotArea>`)
	buf.WriteString(`<c:layout/>`)

	// Generate chart type specific content
	switch opts.ChartKind {
	case ChartKindColumn, ChartKindBar: // both emit <c:barChart> with different barDir
		buf.WriteString(generateBarChartXML(opts))
	case ChartKindLine:
		buf.WriteString(generateLineChartXML(opts))
	case ChartKindPie:
		buf.WriteString(generatePieChartXML(opts))
	case ChartKindArea:
		buf.WriteString(generateAreaChartXML(opts))
	case ChartKindScatter:
		buf.WriteString(generateScatterChartXML(opts))
	default:
		buf.WriteString(generateBarChartXML(opts)) // Default to bar/column
	}

	// Axes (category and value for most chart types, except pie)
	if opts.ChartKind != ChartKindPie {
		buf.WriteString(generateCategoryAxisXML(opts.CategoryAxis))
		buf.WriteString(generateValueAxisXML(opts.ValueAxis))
	}

	buf.WriteString(`</c:plotArea>`)

	// Legend
	if opts.Legend.Show {
		buf.WriteString(generateLegendXML(opts.Legend))
	}

	buf.WriteString(fmt.Sprintf(`<c:plotVisOnly val="%d"/>`, boolToInt(opts.Properties.PlotVisibleOnly)))
	buf.WriteString(fmt.Sprintf(`<c:dispBlanksAs val="%s"/>`, opts.Properties.DisplayBlanksAs))
	buf.WriteString(fmt.Sprintf(`<c:showDLblsOverMax val="%d"/>`, boolToInt(opts.Properties.ShowDataLabelsOverMax)))

	buf.WriteString(`</c:chart>`)

	// External data reference
	buf.WriteString(`<c:externalData r:id="rId1">`)
	buf.WriteString(`<c:autoUpdate val="0"/>`)
	buf.WriteString(`</c:externalData>`)

	buf.WriteString(`</c:chartSpace>`)

	return buf.Bytes()
}

// generateBarChartXML generates bar/column chart XML with extended options
func generateBarChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:barChart>`)
	buf.WriteString(fmt.Sprintf(`<c:barDir val="%s"/>`, opts.BarChartOptions.Direction))
	buf.WriteString(fmt.Sprintf(`<c:grouping val="%s"/>`, opts.BarChartOptions.Grouping))
	buf.WriteString(fmt.Sprintf(`<c:varyColors val="%d"/>`, boolToInt(opts.BarChartOptions.VaryColors)))

	// Series
	for i, series := range opts.Series {
		buf.WriteString(generateSeriesXML(i, series, opts))
	}

	// Data labels (chart-level default)
	if opts.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(opts.DataLabels))
	} else {
		buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`)
	}

	buf.WriteString(fmt.Sprintf(`<c:gapWidth val="%d"/>`, opts.BarChartOptions.GapWidth))
	buf.WriteString(fmt.Sprintf(`<c:overlap val="%d"/>`, opts.BarChartOptions.Overlap))
	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`</c:barChart>`)

	return buf.String()
}

// generateLineChartXML generates line chart XML with extended options
func generateLineChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:lineChart>`)
	buf.WriteString(`<c:grouping val="standard"/>`)
	buf.WriteString(`<c:varyColors val="0"/>`)

	// Series
	for i, series := range opts.Series {
		buf.WriteString(generateSeriesXML(i, series, opts))
	}

	// Data labels
	if opts.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(opts.DataLabels))
	} else {
		buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`)
	}

	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`</c:lineChart>`)

	return buf.String()
}

// generatePieChartXML generates pie chart XML with extended options
func generatePieChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:pieChart>`)
	buf.WriteString(`<c:varyColors val="1"/>`) // Pie charts typically vary colors

	// Series (pie charts usually have one series)
	for i, series := range opts.Series {
		buf.WriteString(generateSeriesXML(i, series, opts))
	}

	// Data labels
	if opts.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(opts.DataLabels))
	} else {
		buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>`)
		buf.WriteString(`<c:showLeaderLines val="1"/></c:dLbls>`)
	}

	buf.WriteString(`</c:pieChart>`)

	return buf.String()
}

// generateAreaChartXML generates area chart XML with extended options
func generateAreaChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:areaChart>`)
	buf.WriteString(`<c:grouping val="standard"/>`)
	buf.WriteString(`<c:varyColors val="0"/>`)

	// Series
	for i, series := range opts.Series {
		buf.WriteString(generateSeriesXML(i, series, opts))
	}

	// Data labels
	if opts.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(opts.DataLabels))
	} else {
		buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`)
	}

	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`</c:areaChart>`)

	return buf.String()
}

// ScatterChartOptions defines options specific to scatter charts
type ScatterChartOptions struct {
	// ScatterStyle defines the scatter chart style
	// "line" - scatter chart with straight lines
	// "lineMarker" - scatter chart with markers and straight lines
	// "marker" - scatter chart with markers only (default)
	// "smooth" - scatter chart with smooth lines
	// "smoothMarker" - scatter chart with markers and smooth lines
	ScatterStyle string

	// VaryColors - whether to vary colors by point
	VaryColors bool
}

// generateScatterChartXML generates scatter chart XML with extended options
func generateScatterChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:scatterChart>`)

	// Scatter style - default to markers only
	scatterStyle := "marker"
	if scOpts := opts.ScatterChartOptions; scOpts != nil {
		if scOpts.ScatterStyle != "" {
			scatterStyle = scOpts.ScatterStyle
		}
	}
	buf.WriteString(fmt.Sprintf(`<c:scatterStyle val="%s"/>`, scatterStyle))

	varyColors := false
	if scOpts := opts.ScatterChartOptions; scOpts != nil {
		varyColors = scOpts.VaryColors
	}
	buf.WriteString(fmt.Sprintf(`<c:varyColors val="%d"/>`, boolToInt(varyColors)))

	// Series - scatter charts use different X values (not categories)
	for i, series := range opts.Series {
		buf.WriteString(generateScatterSeriesXML(i, series, opts))
	}

	// Data labels
	if opts.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(opts.DataLabels))
	} else {
		buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`)
	}

	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`</c:scatterChart>`)

	return buf.String()
}

// generateScatterSeriesXML generates series XML for scatter charts
// Scatter charts use X values instead of categories
func generateScatterSeriesXML(index int, series SeriesOptions, opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`<c:ser><c:idx val="%d"/><c:order val="%d"/>`, index, index))

	// Series name
	buf.WriteString(fmt.Sprintf(`<c:tx><c:strRef><c:f>Sheet1!$%s$1</c:f>`, columnLetter(index+2)))
	buf.WriteString(fmt.Sprintf(`<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>`,
		xmlEscape(series.Name)))

	// Shape properties (color)
	if series.Color != "" {
		buf.WriteString(`<c:spPr>`)
		color := normalizeHexColor(series.Color)
		buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, color))
		buf.WriteString(`</c:spPr>`)
	}

	// X values (instead of categories) - use XValues if provided, otherwise use category indices
	// Column offset: +2 for series data (B,C,D... when no XValues, C,D,E... when XValues exist)
	colOffset := 2
	colLetter := columnLetter(index + colOffset)

	xValColLetter := columnLetter(2) // Column B for X values
	if len(series.XValues) > 0 {
		// X values come from column B in the Excel sheet
		buf.WriteString(fmt.Sprintf(`<c:xVal><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f>`,
			xValColLetter, xValColLetter, len(opts.Categories)+1))
		buf.WriteString(fmt.Sprintf(`<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>`, len(series.XValues)))
		for j, xVal := range series.XValues {
			buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%g</c:v></c:pt>`, j, xVal))
		}
	} else {
		// Fall back to category indices as X values
		buf.WriteString(fmt.Sprintf(`<c:xVal><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f>`,
			colLetter, colLetter, len(opts.Categories)+1))
		buf.WriteString(fmt.Sprintf(`<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>`, len(series.Values)))
		for j := 0; j < len(series.Values) && j < len(opts.Categories); j++ {
			buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%d</c:v></c:pt>`, j, j+1))
		}
	}
	buf.WriteString(`</c:numCache></c:numRef></c:xVal>`)

	// Y values
	buf.WriteString(fmt.Sprintf(`<c:yVal><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f>`,
		colLetter, colLetter, len(opts.Categories)+1))
	buf.WriteString(fmt.Sprintf(`<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>`, len(series.Values)))
	for j, val := range series.Values {
		buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%g</c:v></c:pt>`, j, val))
	}
	buf.WriteString(`</c:numCache></c:numRef></c:yVal>`)

	// Smooth line for scatter
	scatterStyle := "marker"
	if scOpts := opts.ScatterChartOptions; scOpts != nil {
		if strings.HasPrefix(scOpts.ScatterStyle, "smooth") {
			buf.WriteString(`<c:smooth val="1"/>`)
		}
	}

	// Marker for scatter
	if !strings.Contains(scatterStyle, "line") || strings.Contains(scatterStyle, "Marker") {
		buf.WriteString(`<c:marker><c:symbol val="circle"/></c:marker>`)
	} else {
		buf.WriteString(`<c:marker><c:symbol val="none"/></c:marker>`)
	}

	// Per-series data labels
	if series.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(series.DataLabels))
	}

	buf.WriteString(`</c:ser>`)

	return buf.String()
}

// columnLetter converts column number to Excel column letter (1=A, 2=B, etc.)
func columnLetter(col int) string {
	result := ""
	for col > 0 {
		col--
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// createEmbeddedWorkbook creates the embedded Excel workbook with chart data
func (u *Updater) createEmbeddedWorkbook(workbookPath string, opts ChartOptions) error {
	if err := os.MkdirAll(filepath.Dir(workbookPath), 0o755); err != nil {
		return fmt.Errorf("create embeddings directory: %w", err)
	}

	// Create a minimal XLSX file with the chart data
	file, err := os.Create(workbookPath)
	if err != nil {
		return fmt.Errorf("create workbook file: %w", err)
	}
	defer file.Close()

	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// Create [Content_Types].xml
	if err := addZipFile(zipWriter, "[Content_Types].xml", generateWorkbookContentTypes()); err != nil {
		return err
	}

	// Create _rels/.rels
	if err := addZipFile(zipWriter, "_rels/.rels", generateWorkbookRels()); err != nil {
		return err
	}

	// Create xl/workbook.xml
	if err := addZipFile(zipWriter, "xl/workbook.xml", generateWorkbookXML()); err != nil {
		return err
	}

	// Create xl/_rels/workbook.xml.rels
	if err := addZipFile(zipWriter, "xl/_rels/workbook.xml.rels", generateWorkbookXMLRels()); err != nil {
		return err
	}

	// Create xl/worksheets/sheet1.xml with data
	if err := addZipFile(zipWriter, "xl/worksheets/sheet1.xml", generateSheetXML(opts)); err != nil {
		return err
	}

	// Create xl/styles.xml
	if err := addZipFile(zipWriter, "xl/styles.xml", generateStylesXML()); err != nil {
		return err
	}

	return nil
}

// Helper function to add file to zip
func addZipFile(zipWriter *zip.Writer, name string, content []byte) error {
	writer, err := zipWriter.Create(name)
	if err != nil {
		return fmt.Errorf("create zip entry %s: %w", name, err)
	}
	if _, err := writer.Write(content); err != nil {
		return fmt.Errorf("write zip entry %s: %w", name, err)
	}
	return nil
}

// generateWorkbookContentTypes creates the [Content_Types].xml for the embedded workbook
func generateWorkbookContentTypes() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`)
}

// generateWorkbookRels creates the _rels/.rels for the embedded workbook
func generateWorkbookRels() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`)
}

// generateWorkbookXML creates the xl/workbook.xml
func generateWorkbookXML() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`)
}

// generateWorkbookXMLRels creates the xl/_rels/workbook.xml.rels
func generateWorkbookXMLRels() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`)
}

// generateSheetXML creates the xl/worksheets/sheet1.xml with chart data
func generateSheetXML(opts ChartOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>`)

	// Determine if we need X values column (for scatter charts with XValues)
	hasXValues := opts.ChartKind == ChartKindScatter && len(opts.Series) > 0 && len(opts.Series[0].XValues) > 0

	// Header row with series names
	buf.WriteString(`<row r="1">`)
	buf.WriteString(`<c r="A1" t="str"><v></v></c>`) // Empty cell at A1
	if hasXValues {
		buf.WriteString(`<c r="B1" t="str"><v>X Values</v></c>`) // X values header
	}
	for i, series := range opts.Series {
		col := columnLetter(i + 3) // Start from column C if XValues exist, otherwise B
		buf.WriteString(fmt.Sprintf(`<c r="%s1" t="str"><v>%s</v></c>`, col, xmlEscape(series.Name)))
	}
	buf.WriteString(`</row>`)

	// Data rows
	for i, category := range opts.Categories {
		rowNum := i + 2
		buf.WriteString(fmt.Sprintf(`<row r="%d">`, rowNum))

		// Category in column A
		buf.WriteString(fmt.Sprintf(`<c r="A%d" t="str"><v>%s</v></c>`, rowNum, xmlEscape(category)))

		// X values in column B (if applicable)
		if hasXValues && len(opts.Series[0].XValues) > i {
			buf.WriteString(fmt.Sprintf(`<c r="B%d"><v>%g</v></c>`, rowNum, opts.Series[0].XValues[i]))
		}

		// Values for each series
		for j := range opts.Series {
			col := columnLetter(j + 3)
			if i < len(opts.Series[j].Values) {
				buf.WriteString(fmt.Sprintf(`<c r="%s%d"><v>%g</v></c>`, col, rowNum, opts.Series[j].Values[i]))
			}
		}

		buf.WriteString(`</row>`)
	}

	buf.WriteString(`</sheetData>
</worksheet>`)

	return buf.Bytes()
}

// generateStylesXML creates a minimal xl/styles.xml
func generateStylesXML() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="0"/>
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`)
}

// createChartRelationships creates the chart relationships file
func (u *Updater) createChartRelationships(relsPath, workbookPath string) error {
	if err := os.MkdirAll(filepath.Dir(relsPath), 0o755); err != nil {
		return fmt.Errorf("create chart _rels directory: %w", err)
	}

	// Get relative path from charts directory to workbook
	chartsDir := filepath.Join(u.tempDir, "word", "charts")
	relPath, err := filepath.Rel(chartsDir, workbookPath)
	if err != nil {
		return fmt.Errorf("calculate relative path: %w", err)
	}

	// Convert to forward slashes for XML
	relPath = filepath.ToSlash(relPath)

	xml := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="%s"/>
</Relationships>`, relPath)

	if err := atomicWriteFile(relsPath, []byte(xml), 0o644); err != nil {
		return fmt.Errorf("write relationships file: %w", err)
	}

	return nil
}

// insertChartDrawing inserts the chart drawing into the document
func (u *Updater) insertChartDrawing(chartIndex int, relID string, opts ChartOptions) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Generate chart drawing XML
	drawingXML, err := u.generateChartDrawingWithSize(chartIndex, relID, opts.Width, opts.Height)
	if err != nil {
		return fmt.Errorf("generate drawing xml: %w", err)
	}

	// Handle caption if specified
	contentToInsert := drawingXML
	if opts.Caption != nil {
		// Validate caption options
		if err := ValidateCaptionOptions(opts.Caption); err != nil {
			return fmt.Errorf("invalid caption options: %w", err)
		}

		// Set caption type to Figure if not already set
		if opts.Caption.Type == "" {
			opts.Caption.Type = CaptionFigure
		}

		// Generate caption XML
		captionXML := generateCaptionXML(*opts.Caption)

		// Combine chart and caption based on position
		contentToInsert = insertCaptionWithElement(captionXML, drawingXML, opts.Caption.Position)
	}

	// Insert based on position
	var updated []byte
	switch opts.Position {
	case PositionBeginning:
		updated, err = insertAtBodyStart(raw, contentToInsert)
	case PositionEnd:
		updated, err = insertAtBodyEnd(raw, contentToInsert)
	case PositionAfterText:
		if opts.Anchor == "" {
			return fmt.Errorf("anchor text required for PositionAfterText")
		}
		updated, err = insertAfterText(raw, contentToInsert, opts.Anchor)
	case PositionBeforeText:
		if opts.Anchor == "" {
			return fmt.Errorf("anchor text required for PositionBeforeText")
		}
		updated, err = insertBeforeText(raw, contentToInsert, opts.Anchor)
	default:
		return fmt.Errorf("invalid insert position")
	}

	if err != nil {
		return fmt.Errorf("insert chart: %w", err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// generateChartDrawingWithSize creates the inline drawing XML for a chart with custom dimensions
func (u *Updater) generateChartDrawingWithSize(chartIndex int, relId string, width, height int) ([]byte, error) {
	// Get a unique docPr ID (document-wide drawing object ID)
	docPrId, err := u.getNextDocPrId()
	if err != nil {
		return nil, fmt.Errorf("get next docPr id: %w", err)
	}

	// Generate unique IDs
	anchorId := ChartAnchorIDBase + uint32(chartIndex)*ChartIDIncrement
	editId := ChartEditIDBase + uint32(chartIndex)*ChartIDIncrement

	template := `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" wp14:anchorId="%08X" wp14:editId="%08X" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"><wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="15875" b="12700"/><wp:docPr id="%d" name="Chart %d"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

	return fmt.Appendf(nil, template, anchorId, editId, width, height, docPrId, chartIndex, relId), nil
}

// applyAxisDefaults applies default values to axis options
func applyAxisDefaults(axis *AxisOptions, isCategoryAxis bool) *AxisOptions {
	if axis == nil {
		axis = &AxisOptions{}
	}

	axis.Visible = true // Always visible by default

	if axis.Position == "" {
		if isCategoryAxis {
			axis.Position = AxisPositionBottom
		} else {
			axis.Position = AxisPositionLeft
		}
	}

	if axis.MajorTickMark == "" {
		axis.MajorTickMark = TickMarkOut
	}
	if axis.MinorTickMark == "" {
		axis.MinorTickMark = TickMarkNone
	}
	if axis.TickLabelPos == "" {
		axis.TickLabelPos = TickLabelNextTo
	}
	if axis.NumberFormat == "" {
		axis.NumberFormat = "General"
	}

	// Value axis has major gridlines by default
	if !isCategoryAxis {
		axis.MajorGridlines = true
	}

	return axis
}

// generateTitleXML generates chart title XML
func generateTitleXML(title string, overlay bool) string {
	var buf bytes.Buffer
	buf.WriteString(`<c:title>`)
	buf.WriteString(`<c:tx>`)
	buf.WriteString(`<c:rich>`)
	buf.WriteString(`<a:bodyPr/>`)
	buf.WriteString(`<a:lstStyle/>`)
	buf.WriteString(`<a:p>`)
	buf.WriteString(`<a:pPr><a:defRPr/></a:pPr>`)
	buf.WriteString(`<a:r><a:rPr lang="en-US"/><a:t>`)
	buf.WriteString(xmlEscape(title))
	buf.WriteString(`</a:t></a:r>`)
	buf.WriteString(`</a:p>`)
	buf.WriteString(`</c:rich>`)
	buf.WriteString(`</c:tx>`)
	buf.WriteString(`<c:layout/>`)
	buf.WriteString(fmt.Sprintf(`<c:overlay val="%d"/>`, boolToInt(overlay)))
	buf.WriteString(`</c:title>`)
	return buf.String()
}

// generateLegendXML generates legend XML
func generateLegendXML(legend *LegendOptions) string {
	var buf bytes.Buffer
	buf.WriteString(`<c:legend>`)
	buf.WriteString(fmt.Sprintf(`<c:legendPos val="%s"/>`, legend.Position))
	buf.WriteString(`<c:layout/>`)
	buf.WriteString(fmt.Sprintf(`<c:overlay val="%d"/>`, boolToInt(legend.Overlay)))
	buf.WriteString(`</c:legend>`)
	return buf.String()
}

// generateSeriesXML generates series XML
func generateSeriesXML(index int, series SeriesOptions, opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`<c:ser><c:idx val="%d"/><c:order val="%d"/>`, index, index))

	// Series name
	buf.WriteString(fmt.Sprintf(`<c:tx><c:strRef><c:f>Sheet1!$%s$1</c:f>`, columnLetter(index+2)))
	buf.WriteString(fmt.Sprintf(`<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>`,
		xmlEscape(series.Name)))

	// Shape properties (color, etc.)
	if series.Color != "" || series.InvertIfNegative {
		buf.WriteString(`<c:spPr>`)
		if series.Color != "" {
			color := normalizeHexColor(series.Color)
			buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, color))
		}
		buf.WriteString(`</c:spPr>`)
	}

	// Invert if negative
	if series.InvertIfNegative {
		buf.WriteString(`<c:invertIfNegative val="1"/>`)
	}

	// Categories
	buf.WriteString(fmt.Sprintf(`<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$%d</c:f>`, len(opts.Categories)+1))
	buf.WriteString(fmt.Sprintf(`<c:strCache><c:ptCount val="%d"/>`, len(opts.Categories)))
	for j, cat := range opts.Categories {
		buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%s</c:v></c:pt>`, j, xmlEscape(cat)))
	}
	buf.WriteString(`</c:strCache></c:strRef></c:cat>`)

	// Values
	colLetter := columnLetter(index + 2)
	buf.WriteString(fmt.Sprintf(`<c:val><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f>`,
		colLetter, colLetter, len(opts.Categories)+1))
	buf.WriteString(fmt.Sprintf(`<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>`, len(series.Values)))
	for j, val := range series.Values {
		buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%g</c:v></c:pt>`, j, val))
	}
	buf.WriteString(`</c:numCache></c:numRef></c:val>`)

	// Line chart specific: smooth and markers
	if opts.ChartKind == ChartKindLine {
		if series.Smooth {
			buf.WriteString(`<c:smooth val="1"/>`)
		}
		if series.ShowMarkers {
			buf.WriteString(`<c:marker><c:symbol val="circle"/></c:marker>`)
		} else {
			buf.WriteString(`<c:marker><c:symbol val="none"/></c:marker>`)
		}
	}

	// Per-series data labels (overrides chart-level)
	if series.DataLabels != nil {
		buf.WriteString(generateDataLabelsXML(series.DataLabels))
	}

	buf.WriteString(`</c:ser>`)

	return buf.String()
}

// generateDataLabelsXML generates data labels XML
func generateDataLabelsXML(labels *DataLabelOptions) string {
	var buf bytes.Buffer
	buf.WriteString(`<c:dLbls>`)
	buf.WriteString(fmt.Sprintf(`<c:showLegendKey val="%d"/>`, boolToInt(labels.ShowLegendKey)))
	buf.WriteString(fmt.Sprintf(`<c:showVal val="%d"/>`, boolToInt(labels.ShowValue)))
	buf.WriteString(fmt.Sprintf(`<c:showCatName val="%d"/>`, boolToInt(labels.ShowCategoryName)))
	buf.WriteString(fmt.Sprintf(`<c:showSerName val="%d"/>`, boolToInt(labels.ShowSeriesName)))
	buf.WriteString(fmt.Sprintf(`<c:showPercent val="%d"/>`, boolToInt(labels.ShowPercent)))
	buf.WriteString(`<c:showBubbleSize val="0"/>`)
	if labels.Position != "" {
		buf.WriteString(fmt.Sprintf(`<c:dLblPos val="%s"/>`, labels.Position))
	}
	if labels.ShowLeaderLines {
		buf.WriteString(`<c:showLeaderLines val="1"/>`)
	}
	buf.WriteString(`</c:dLbls>`)
	return buf.String()
}

// generateCategoryAxisXML generates category axis XML with extended options
func generateCategoryAxisXML(axis *AxisOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:catAx>`)
	buf.WriteString(`<c:axId val="2071991400"/>`)

	// Scaling
	buf.WriteString(`<c:scaling>`)
	if axis.Min != nil {
		buf.WriteString(fmt.Sprintf(`<c:min val="%g"/>`, *axis.Min))
	}
	if axis.Max != nil {
		buf.WriteString(fmt.Sprintf(`<c:max val="%g"/>`, *axis.Max))
	}
	buf.WriteString(`<c:orientation val="minMax"/>`)
	buf.WriteString(`</c:scaling>`)

	buf.WriteString(fmt.Sprintf(`<c:delete val="%d"/>`, boolToInt(!axis.Visible)))
	buf.WriteString(fmt.Sprintf(`<c:axPos val="%s"/>`, axis.Position))

	// Major gridlines
	if axis.MajorGridlines {
		buf.WriteString(`<c:majorGridlines/>`)
	}

	// Title
	if axis.Title != "" {
		buf.WriteString(generateAxisTitleXML(axis.Title, axis.TitleOverlay))
	}

	buf.WriteString(fmt.Sprintf(`<c:numFmt formatCode="%s" sourceLinked="0"/>`, axis.NumberFormat))
	buf.WriteString(fmt.Sprintf(`<c:majorTickMark val="%s"/>`, axis.MajorTickMark))
	buf.WriteString(fmt.Sprintf(`<c:minorTickMark val="%s"/>`, axis.MinorTickMark))
	buf.WriteString(fmt.Sprintf(`<c:tickLblPos val="%s"/>`, axis.TickLabelPos))

	buf.WriteString(`<c:crossAx val="2071991240"/>`)

	if axis.CrossesAt != nil {
		buf.WriteString(fmt.Sprintf(`<c:crossesAt val="%g"/>`, *axis.CrossesAt))
	} else {
		buf.WriteString(`<c:crosses val="autoZero"/>`)
	}

	buf.WriteString(`<c:auto val="1"/>`)
	buf.WriteString(`<c:lblAlgn val="ctr"/>`)
	buf.WriteString(`<c:lblOffset val="100"/>`)

	// Minor gridlines
	if axis.MinorGridlines {
		buf.WriteString(`<c:minorGridlines/>`)
	}

	buf.WriteString(`</c:catAx>`)

	return buf.String()
}

// generateValueAxisXML generates value axis XML with extended options
func generateValueAxisXML(axis *AxisOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:valAx>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)

	// Scaling
	buf.WriteString(`<c:scaling>`)
	if axis.Min != nil {
		buf.WriteString(fmt.Sprintf(`<c:min val="%g"/>`, *axis.Min))
	}
	if axis.Max != nil {
		buf.WriteString(fmt.Sprintf(`<c:max val="%g"/>`, *axis.Max))
	}
	buf.WriteString(`<c:orientation val="minMax"/>`)
	buf.WriteString(`</c:scaling>`)

	buf.WriteString(fmt.Sprintf(`<c:delete val="%d"/>`, boolToInt(!axis.Visible)))
	buf.WriteString(fmt.Sprintf(`<c:axPos val="%s"/>`, axis.Position))

	// Major gridlines
	if axis.MajorGridlines {
		buf.WriteString(`<c:majorGridlines/>`)
	}

	// Title
	if axis.Title != "" {
		buf.WriteString(generateAxisTitleXML(axis.Title, axis.TitleOverlay))
	}

	buf.WriteString(fmt.Sprintf(`<c:numFmt formatCode="%s" sourceLinked="0"/>`, axis.NumberFormat))

	if axis.MajorUnit != nil {
		buf.WriteString(fmt.Sprintf(`<c:majorUnit val="%g"/>`, *axis.MajorUnit))
	}
	if axis.MinorUnit != nil {
		buf.WriteString(fmt.Sprintf(`<c:minorUnit val="%g"/>`, *axis.MinorUnit))
	}

	buf.WriteString(fmt.Sprintf(`<c:majorTickMark val="%s"/>`, axis.MajorTickMark))
	buf.WriteString(fmt.Sprintf(`<c:minorTickMark val="%s"/>`, axis.MinorTickMark))
	buf.WriteString(fmt.Sprintf(`<c:tickLblPos val="%s"/>`, axis.TickLabelPos))

	buf.WriteString(`<c:crossAx val="2071991400"/>`)

	if axis.CrossesAt != nil {
		buf.WriteString(fmt.Sprintf(`<c:crossesAt val="%g"/>`, *axis.CrossesAt))
	} else {
		buf.WriteString(`<c:crosses val="autoZero"/>`)
	}

	buf.WriteString(`<c:crossBetween val="between"/>`)

	// Minor gridlines
	if axis.MinorGridlines {
		buf.WriteString(`<c:minorGridlines/>`)
	}

	buf.WriteString(`</c:valAx>`)

	return buf.String()
}

// generateAxisTitleXML generates axis title XML
func generateAxisTitleXML(title string, overlay bool) string {
	var buf bytes.Buffer
	buf.WriteString(`<c:title>`)
	buf.WriteString(`<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>`)
	buf.WriteString(`<a:pPr><a:defRPr/></a:pPr>`)
	buf.WriteString(`<a:r><a:rPr lang="en-US"/><a:t>`)
	buf.WriteString(xmlEscape(title))
	buf.WriteString(`</a:t></a:r>`)
	buf.WriteString(`</a:p></c:rich></c:tx>`)
	buf.WriteString(`<c:layout/>`)
	buf.WriteString(fmt.Sprintf(`<c:overlay val="%d"/>`, boolToInt(overlay)))
	buf.WriteString(`</c:title>`)
	return buf.String()
}

// boolToInt converts boolean to integer (0 or 1)
func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

// findNextChartIndex finds the next available chart index by scanning chart files.
func (u *Updater) findNextChartIndex() int {
	chartsDir := filepath.Join(u.tempDir, "word", "charts")
	entries, err := os.ReadDir(chartsDir)
	if err != nil {
		return 1
	}

	maxIndex := 0
	for _, entry := range entries {
		if matches := chartFilePattern.FindStringSubmatch(entry.Name()); matches != nil {
			idx, err := strconv.Atoi(matches[1])
			if err != nil {
				continue
			}
			if idx > maxIndex {
				maxIndex = idx
			}
		}
	}

	return maxIndex + 1
}

// addChartRelationship appends a Relationship for a chart to document.xml.rels and returns its Id.
func (u *Updater) addChartRelationship(chartIndex int) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document relationships: %w", err)
	}

	nextRelId, err := u.getNextDocumentRelId()
	if err != nil {
		return "", err
	}

	insert := fmt.Sprintf("\n  <Relationship Id=\"%s\" Type=\"%s/chart\" Target=\"charts/chart%d.xml\"/>\n", nextRelId, OfficeDocumentNS, chartIndex)
	closer := []byte("</Relationships>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return "", fmt.Errorf("invalid document.xml.rels: missing </Relationships>")
	}
	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])

	if err := atomicWriteFile(relsPath, result, 0o644); err != nil {
		return "", fmt.Errorf("write relationships: %w", err)
	}
	return nextRelId, nil
}

// addContentTypeOverride adds a content type override for a chart in [Content_Types].xml.
func (u *Updater) addContentTypeOverride(chartIndex int) error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	raw, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	chartPart := fmt.Sprintf("/word/charts/chart%d.xml", chartIndex)
	if bytes.Contains(raw, []byte(chartPart)) {
		return nil // already present
	}

	insert := fmt.Sprintf("\n  <Override PartName=\"%s\" ContentType=\"%s\"/>\n", chartPart, ChartContentType)
	closer := []byte("</Types>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return fmt.Errorf("invalid [Content_Types].xml: missing </Types>")
	}
	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])
	return atomicWriteFile(contentTypesPath, result, 0o644)
}

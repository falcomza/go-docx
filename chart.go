package docxupdater

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// ChartKind defines the type of chart
type ChartKind string

const (
	ChartKindColumn ChartKind = "barChart"  // Column chart (vertical bars)
	ChartKindBar    ChartKind = "barChart"  // Bar chart (horizontal bars)
	ChartKindLine   ChartKind = "lineChart" // Line chart
	ChartKindPie    ChartKind = "pieChart"  // Pie chart
	ChartKindArea   ChartKind = "areaChart" // Area chart
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
	CategoryAxisTitle string // X-axis title (horizontal axis)
	ValueAxisTitle    string // Y-axis title (vertical axis)

	// Data
	Categories     []string     // Category labels (X-axis)
	Series         []SeriesData // Data series with names and values
	ShowLegend     bool         // Show legend (default: true)
	LegendPosition string       // Legend position: "r" (right), "l" (left), "t" (top), "b" (bottom)

	// Chart dimensions (default: spans between margins)
	Width  int // Width in EMUs (English Metric Units), 0 for default (6099523 = ~6.5")
	Height int // Height in EMUs, 0 for default (3340467 = ~3.5")

	// Caption options (nil for no caption)
	Caption *CaptionOptions
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

	return nil
}

// applyChartDefaults sets default values for unspecified options
func applyChartDefaults(opts ChartOptions) ChartOptions {
	if opts.ChartKind == "" {
		opts.ChartKind = ChartKindColumn
	}
	if opts.Width == 0 {
		opts.Width = 6099523 // ~6.5 inches (spans between margins on letter-size page)
	}
	if opts.Height == 0 {
		opts.Height = 3340467 // ~3.5 inches
	}
	if opts.ShowLegend && opts.LegendPosition == "" {
		opts.LegendPosition = "r" // Right by default
	}
	return opts
}

// createChartXML generates the chart XML file
func (u *Updater) createChartXML(chartPath string, opts ChartOptions) error {
	if err := os.MkdirAll(filepath.Dir(chartPath), 0o755); err != nil {
		return fmt.Errorf("create charts directory: %w", err)
	}

	xml := generateChartXML(opts)

	if err := os.WriteFile(chartPath, xml, 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

// generateChartXML creates the chart XML content
func generateChartXML(opts ChartOptions) []byte {
	var buf bytes.Buffer

	// XML declaration with newline (Word requires this)
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")

	// chartSpace with all standard namespaces (Word requires these for compatibility)
	buf.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`)
	buf.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">`)

	// Chart properties (Word expects these even if not modified)
	buf.WriteString(`<c:date1904 val="0"/>`)
	buf.WriteString(`<c:lang val="en-US"/>`)
	buf.WriteString(`<c:roundedCorners val="0"/>`)

	buf.WriteString(`<c:chart>`)

	// Chart title
	if opts.Title != "" {
		buf.WriteString(`<c:title>`)
		buf.WriteString(`<c:tx>`)
		buf.WriteString(`<c:rich>`)
		buf.WriteString(`<a:bodyPr/>`)
		buf.WriteString(`<a:lstStyle/>`)
		buf.WriteString(`<a:p>`)
		buf.WriteString(`<a:pPr><a:defRPr/></a:pPr>`)
		buf.WriteString(`<a:r><a:rPr lang="en-US"/><a:t>`)
		buf.WriteString(xmlEscape(opts.Title))
		buf.WriteString(`</a:t></a:r>`)
		buf.WriteString(`</a:p>`)
		buf.WriteString(`</c:rich>`)
		buf.WriteString(`</c:tx>`)
		buf.WriteString(`<c:layout/>`)
		buf.WriteString(`<c:overlay val="0"/>`)
		buf.WriteString(`</c:title>`)
	}

	buf.WriteString(`<c:autoTitleDeleted val="0"/>`)
	buf.WriteString(`<c:plotArea>`)
	buf.WriteString(`<c:layout/>`)

	// Generate chart type specific content
	switch opts.ChartKind {
	case ChartKindColumn:
		buf.WriteString(generateColumnChartXML(opts))
	default:
		buf.WriteString(generateColumnChartXML(opts)) // Default to column
	}

	// Category axis
	buf.WriteString(`<c:catAx>`)
	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:scaling><c:orientation val="minMax"/></c:scaling>`)
	buf.WriteString(`<c:delete val="0"/>`)
	buf.WriteString(`<c:axPos val="b"/>`)
	if opts.CategoryAxisTitle != "" {
		buf.WriteString(`<c:title>`)
		buf.WriteString(`<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>`)
		buf.WriteString(xmlEscape(opts.CategoryAxisTitle))
		buf.WriteString(`</a:t></a:r></a:p></c:rich></c:tx>`)
		buf.WriteString(`<c:layout/><c:overlay val="0"/>`)
		buf.WriteString(`</c:title>`)
	}
	buf.WriteString(`<c:numFmt formatCode="General" sourceLinked="1"/>`)
	buf.WriteString(`<c:majorTickMark val="out"/>`)
	buf.WriteString(`<c:minorTickMark val="none"/>`)
	buf.WriteString(`<c:tickLblPos val="nextTo"/>`)
	buf.WriteString(`<c:crossAx val="2071991240"/>`)
	buf.WriteString(`<c:crosses val="autoZero"/>`)
	buf.WriteString(`<c:auto val="1"/>`)
	buf.WriteString(`<c:lblAlgn val="ctr"/>`)
	buf.WriteString(`<c:lblOffset val="100"/>`)
	buf.WriteString(`</c:catAx>`)

	// Value axis
	buf.WriteString(`<c:valAx>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`<c:scaling><c:orientation val="minMax"/></c:scaling>`)
	buf.WriteString(`<c:delete val="0"/>`)
	buf.WriteString(`<c:axPos val="l"/>`)
	if opts.ValueAxisTitle != "" {
		buf.WriteString(`<c:title>`)
		buf.WriteString(`<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>`)
		buf.WriteString(xmlEscape(opts.ValueAxisTitle))
		buf.WriteString(`</a:t></a:r></a:p></c:rich></c:tx>`)
		buf.WriteString(`<c:layout/><c:overlay val="0"/>`)
		buf.WriteString(`</c:title>`)
	}
	buf.WriteString(`<c:numFmt formatCode="General" sourceLinked="1"/>`)
	buf.WriteString(`<c:majorTickMark val="out"/>`)
	buf.WriteString(`<c:minorTickMark val="none"/>`)
	buf.WriteString(`<c:tickLblPos val="nextTo"/>`)
	buf.WriteString(`<c:crossAx val="2071991400"/>`)
	buf.WriteString(`<c:crosses val="autoZero"/>`)
	buf.WriteString(`<c:crossBetween val="between"/>`)
	buf.WriteString(`</c:valAx>`)

	buf.WriteString(`</c:plotArea>`)

	// Legend
	if opts.ShowLegend {
		buf.WriteString(`<c:legend>`)
		buf.WriteString(fmt.Sprintf(`<c:legendPos val="%s"/>`, opts.LegendPosition))
		buf.WriteString(`<c:layout/>`)
		buf.WriteString(`<c:overlay val="0"/>`)
		buf.WriteString(`</c:legend>`)
	}

	buf.WriteString(`<c:plotVisOnly val="1"/>`)
	buf.WriteString(`<c:dispBlanksAs val="gap"/>`)
	buf.WriteString(`<c:showDLblsOverMax val="0"/>`)

	buf.WriteString(`</c:chart>`)

	// External data reference
	buf.WriteString(`<c:externalData r:id="rId1">`)
	buf.WriteString(`<c:autoUpdate val="0"/>`)
	buf.WriteString(`</c:externalData>`)

	buf.WriteString(`</c:chartSpace>`)

	return buf.Bytes()
}

// generateColumnChartXML generates column chart specific XML
func generateColumnChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:barChart>`)
	buf.WriteString(`<c:barDir val="col"/>`) // Column direction (col=vertical, bar=horizontal)
	buf.WriteString(`<c:grouping val="clustered"/>`)
	buf.WriteString(`<c:varyColors val="0"/>`)

	// Series
	for i, series := range opts.Series {
		buf.WriteString(fmt.Sprintf(`<c:ser>
<c:idx val="%d"/>
<c:order val="%d"/>
<c:tx>
  <c:strRef>
    <c:f>Sheet1!$%s$1</c:f>
    <c:strCache>
      <c:ptCount val="1"/>
      <c:pt idx="0"><c:v>%s</c:v></c:pt>
    </c:strCache>
  </c:strRef>
</c:tx>`, i, i, columnLetter(i+1), xmlEscape(series.Name)))

		buf.WriteString(`<c:cat>
  <c:strRef>
    <c:f>Sheet1!$A$2:$A$`)
		buf.WriteString(fmt.Sprintf("%d", len(opts.Categories)+1))
		buf.WriteString(`</c:f>
    <c:strCache>
      <c:ptCount val="`)
		buf.WriteString(fmt.Sprintf("%d", len(opts.Categories)))
		buf.WriteString(`"/>`)
		for j, cat := range opts.Categories {
			buf.WriteString(fmt.Sprintf(`<c:pt  idx="%d"><c:v>%s</c:v></c:pt>`, j, xmlEscape(cat)))
		}
		buf.WriteString(`</c:strCache>
  </c:strRef>
</c:cat>`)

		buf.WriteString(`<c:val>
  <c:numRef>
    <c:f>Sheet1!$`)
		buf.WriteString(columnLetter(i + 1))
		buf.WriteString(`$2:$`)
		buf.WriteString(columnLetter(i + 1))
		buf.WriteString(`$`)
		buf.WriteString(fmt.Sprintf("%d", len(opts.Categories)+1))
		buf.WriteString(`</c:f>
    <c:numCache>
      <c:formatCode>General</c:formatCode>
      <c:ptCount val="`)
		buf.WriteString(fmt.Sprintf("%d", len(series.Values)))
		buf.WriteString(`"/>`)
		for j, val := range series.Values {
			buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%g</c:v></c:pt>`, j, val))
		}
		buf.WriteString(`</c:numCache>
  </c:numRef>
</c:val>`)

		// Add color if specified
		if color := normalizeHexColor(series.Color); color != "" {
			buf.WriteString(`<c:spPr><a:solidFill><a:srgbClr val="`)
			buf.WriteString(color)
			buf.WriteString(`"/></a:solidFill></c:spPr>`)
		}

		buf.WriteString(`</c:ser>`)
	}

	buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`)
	buf.WriteString(`<c:gapWidth val="150"/>`)
	buf.WriteString(`<c:overlap val="0"/>`)
	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`</c:barChart>`)

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

	// Header row with series names
	buf.WriteString(`<row r="1">`)
	buf.WriteString(`<c r="A1" t="str"><v></v></c>`) // Empty cell at A1
	for i, series := range opts.Series {
		col := columnLetter(i + 2) // B, C, D, etc.
		buf.WriteString(fmt.Sprintf(`<c r="%s1" t="str"><v>%s</v></c>`, col, xmlEscape(series.Name)))
	}
	buf.WriteString(`</row>`)

	// Data rows
	for i, category := range opts.Categories {
		rowNum := i + 2
		buf.WriteString(fmt.Sprintf(`<row r="%d">`, rowNum))

		// Category in column A
		buf.WriteString(fmt.Sprintf(`<c r="A%d" t="str"><v>%s</v></c>`, rowNum, xmlEscape(category)))

		// Values for each series
		for j, series := range opts.Series {
			col := columnLetter(j + 2)
			buf.WriteString(fmt.Sprintf(`<c r="%s%d"><v>%g</v></c>`, col, rowNum, series.Values[i]))
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

	if err := os.WriteFile(relsPath, []byte(xml), 0o644); err != nil {
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
		contentToInsert = insertCaptionWithElement(raw, captionXML, drawingXML, opts.Caption.Position)
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

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
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

	template := `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="%08X" wp14:editId="%08X"><wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="15875" b="12700"/><wp:docPr id="%d" name="Chart %d"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

	return fmt.Appendf(nil, template, anchorId, editId, width, height, docPrId, chartIndex, relId), nil
}

// ==================== Extended Chart Functionality ====================

// InsertChartExtended creates a chart with comprehensive customization options
func (u *Updater) InsertChartExtended(opts ExtendedChartOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Validate options
	if err := validateExtendedChartOptions(opts); err != nil {
		return fmt.Errorf("invalid extended chart options: %w", err)
	}

	// Apply defaults
	opts = applyExtendedChartDefaults(opts)

	// Find next available chart index
	chartIndex := u.findNextChartIndex()

	// Create chart XML file
	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	if err := u.createExtendedChartXML(chartPath, opts); err != nil {
		return fmt.Errorf("create chart xml: %w", err)
	}

	// Create embedded workbook (convert to simple format)
	workbookOpts := convertToChartOptions(opts)
	workbookPath := filepath.Join(u.tempDir, "word", "embeddings", fmt.Sprintf("Microsoft_Excel_Worksheet%d.xlsx", chartIndex))
	if err := u.createEmbeddedWorkbook(workbookPath, workbookOpts); err != nil {
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
	if err := u.insertExtendedChartDrawing(chartIndex, relID, opts); err != nil {
		return fmt.Errorf("insert chart drawing: %w", err)
	}

	// Update content types
	if err := u.addContentTypeOverride(chartIndex); err != nil {
		return fmt.Errorf("add content type: %w", err)
	}

	return nil
}

// validateExtendedChartOptions validates extended chart options
func validateExtendedChartOptions(opts ExtendedChartOptions) error {
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
			return fmt.Errorf("series[%d] values length (%d) must match categories length (%d)",
				i, len(series.Values), len(opts.Categories))
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

// applyExtendedChartDefaults applies default values to extended chart options
func applyExtendedChartDefaults(opts ExtendedChartOptions) ExtendedChartOptions {
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

	// Apply legend defaults
	if opts.Legend == nil {
		opts.Legend = &LegendOptions{
			Show:     true,
			Position: "r",
			Overlay:  false,
		}
	}

	// Apply category axis defaults
	if opts.CategoryAxis == nil {
		opts.CategoryAxis = &AxisOptions{}
	}
	opts.CategoryAxis = applyAxisDefaults(opts.CategoryAxis, true)

	// Apply value axis defaults
	if opts.ValueAxis == nil {
		opts.ValueAxis = &AxisOptions{}
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

// convertToChartOptions converts ExtendedChartOptions to ChartOptions for workbook creation
func convertToChartOptions(opts ExtendedChartOptions) ChartOptions {
	// Convert SeriesOptions to SeriesData
	series := make([]SeriesData, len(opts.Series))
	for i, s := range opts.Series {
		series[i] = SeriesData{
			Name:   s.Name,
			Values: s.Values,
			Color:  s.Color,
		}
	}

	return ChartOptions{
		Position:   opts.Position,
		Anchor:     opts.Anchor,
		ChartKind:  opts.ChartKind,
		Title:      opts.Title,
		Categories: opts.Categories,
		Series:     series,
		ShowLegend: opts.Legend != nil && opts.Legend.Show,
		LegendPosition: func() string {
			if opts.Legend != nil {
				return opts.Legend.Position
			}
			return "r"
		}(),
		Width:   opts.Width,
		Height:  opts.Height,
		Caption: opts.Caption,
	}
}

// createExtendedChartXML generates the chart XML file with extended options
func (u *Updater) createExtendedChartXML(chartPath string, opts ExtendedChartOptions) error {
	if err := os.MkdirAll(filepath.Dir(chartPath), 0o755); err != nil {
		return fmt.Errorf("create charts directory: %w", err)
	}

	xml := generateExtendedChartXML(opts)

	if err := os.WriteFile(chartPath, xml, 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

// generateExtendedChartXML creates the chart XML content with all extended options
func generateExtendedChartXML(opts ExtendedChartOptions) []byte {
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
	case "barChart": // ChartKindColumn and ChartKindBar both use barChart
		buf.WriteString(generateExtendedBarChartXML(opts))
	case ChartKindLine:
		buf.WriteString(generateExtendedLineChartXML(opts))
	case ChartKindPie:
		buf.WriteString(generateExtendedPieChartXML(opts))
	case ChartKindArea:
		buf.WriteString(generateExtendedAreaChartXML(opts))
	default:
		buf.WriteString(generateExtendedBarChartXML(opts)) // Default to bar/column
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

// generateExtendedBarChartXML generates bar/column chart XML with extended options
func generateExtendedBarChartXML(opts ExtendedChartOptions) string {
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

// generateExtendedLineChartXML generates line chart XML with extended options
func generateExtendedLineChartXML(opts ExtendedChartOptions) string {
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

// generateExtendedPieChartXML generates pie chart XML with extended options
func generateExtendedPieChartXML(opts ExtendedChartOptions) string {
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

// generateExtendedAreaChartXML generates area chart XML with extended options
func generateExtendedAreaChartXML(opts ExtendedChartOptions) string {
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

// generateSeriesXML generates series XML with extended options
func generateSeriesXML(index int, series SeriesOptions, opts ExtendedChartOptions) string {
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

// insertExtendedChartDrawing inserts the chart drawing into the document
func (u *Updater) insertExtendedChartDrawing(chartIndex int, relId string, opts ExtendedChartOptions) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	drawing, err := u.generateChartDrawingWithSize(chartIndex, relId, opts.Width, opts.Height)
	if err != nil {
		return fmt.Errorf("generate chart drawing: %w", err)
	}

	// Add caption if specified
	contentToInsert := drawing
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
		contentToInsert = insertCaptionWithElement(raw, captionXML, drawing, opts.Caption.Position)
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

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// boolToInt converts boolean to integer (0 or 1)
func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

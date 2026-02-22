package godocx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"strconv"
	"strings"
)

// ChartSpace represents the root element of a chart XML document.
type ChartSpace struct {
	XMLName      xml.Name `xml:"chartSpace"`
	Chart        Chart    `xml:"chart"`
	ExternalData struct {
		RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	} `xml:"externalData"`
}

func parseScatterXValues(categories []string) ([]float64, error) {
	values := make([]float64, len(categories))
	for i, cat := range categories {
		clean := strings.TrimSpace(cat)
		if clean == "" {
			return nil, fmt.Errorf("scatter chart categories must be numeric")
		}
		val, err := strconv.ParseFloat(clean, 64)
		if err != nil {
			return nil, fmt.Errorf("scatter chart categories must be numeric: %w", err)
		}
		values[i] = val
	}
	return values, nil
}

// Chart represents the chart element containing plot area.
type Chart struct {
	Title    *ChartTitle `xml:"title,omitempty"`
	PlotArea PlotArea    `xml:"plotArea"`
}

// ChartTitle represents the chart title element.
type ChartTitle struct {
	Tx struct {
		Rich *struct {
			P struct {
				R struct {
					T string `xml:"t"`
				} `xml:"r"`
			} `xml:"p"`
		} `xml:"rich,omitempty"`
		StrRef *struct {
			F string `xml:"f"`
		} `xml:"strRef,omitempty"`
	} `xml:"tx"`
}

// PlotArea contains the chart type and series.
type PlotArea struct {
	BarChart     *ChartType `xml:"barChart,omitempty"`
	LineChart    *ChartType `xml:"lineChart,omitempty"`
	ScatterChart *ChartType `xml:"scatterChart,omitempty"`
	PieChart     *ChartType `xml:"pieChart,omitempty"`
	CatAx        []Axis     `xml:"catAx,omitempty"`
	ValAx        []Axis     `xml:"valAx,omitempty"`
}

// ChartType represents a chart with series data.
type ChartType struct {
	Series []Series `xml:"ser"`
}

// Series represents a data series in a chart.
type Series struct {
	Tx  SeriesText `xml:"tx"`
	Cat *SeriesRef `xml:"cat,omitempty"`
	Val *SeriesRef `xml:"val,omitempty"`
	// For scatter charts
	XVal *SeriesRef `xml:"xVal,omitempty"`
	YVal *SeriesRef `xml:"yVal,omitempty"`
}

// SeriesText contains the series name.
type SeriesText struct {
	V   string     `xml:"v,omitempty"`
	Ref *SeriesRef `xml:"strRef,omitempty"`
}

// SeriesRef represents a reference to chart data.
type SeriesRef struct {
	F        string    `xml:"f,omitempty"`
	StrCache *StrCache `xml:"strCache,omitempty"`
	NumCache *NumCache `xml:"numCache,omitempty"`
}

// StrCache contains cached string values.
type StrCache struct {
	PtCount int     `xml:"ptCount>val,attr"`
	Pts     []StrPt `xml:"pt"`
}

// StrPt represents a string point in the cache.
type StrPt struct {
	Idx int    `xml:"idx,attr"`
	V   string `xml:"v"`
}

// NumCache contains cached numeric values.
type NumCache struct {
	PtCount int     `xml:"ptCount>val,attr"`
	Pts     []NumPt `xml:"pt"`
}

// NumPt represents a numeric point in the cache.
type NumPt struct {
	Idx int     `xml:"idx,attr"`
	V   float64 `xml:"v"`
}

// Axis represents a chart axis.
type Axis struct {
	Title *AxisTitle `xml:"title,omitempty"`
}

// AxisTitle represents an axis title.
type AxisTitle struct {
	Tx struct {
		Rich *struct {
			P struct {
				R struct {
					T string `xml:"t"`
				} `xml:"r"`
			} `xml:"p"`
		} `xml:"rich,omitempty"`
	} `xml:"tx"`
}

// updateChartXML updates the chart XML file with new data.
func updateChartXML(chartPath string, data ChartData) error {
	rawXML, err := os.ReadFile(chartPath)
	if err != nil {
		return fmt.Errorf("read chart xml: %w", err)
	}

	// Use direct XML manipulation for better namespace support
	// Parse to understand structure, then manipulate the raw XML
	updated, err := updateChartXMLContent(rawXML, data)
	if err != nil {
		return fmt.Errorf("update chart content: %w", err)
	}

	// Ensure proper XML formatting: verify newline after XML declaration
	updated = ensureXMLDeclarationNewline(updated)

	if err := os.WriteFile(chartPath, updated, 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

// ensureXMLDeclarationNewline ensures there's a newline after the XML declaration
// Word requires this for strict XML parsing
func ensureXMLDeclarationNewline(xmlContent []byte) []byte {
	content := string(xmlContent)

	// Find the XML declaration
	declEnd := "?>"
	idx := strings.Index(content, declEnd)
	if idx == -1 {
		return xmlContent // No XML declaration found
	}

	idx += len(declEnd)

	// Check if there's already a newline
	if idx < len(content) && content[idx] == '\n' {
		return xmlContent // Already has newline
	}

	// Insert newline after XML declaration
	result := content[:idx] + "\n" + content[idx:]
	return []byte(result)
}

// updateChartXMLContent updates chart XML content handling multiple namespaces and chart types.
func updateChartXMLContent(rawXML []byte, data ChartData) ([]byte, error) {
	content := string(rawXML)

	// Detect namespace prefix (c: or no prefix)
	nsPrefix := detectNamespacePrefix(content)

	// Update title if provided
	if data.ChartTitle != "" {
		content = updateChartTitle(content, data.ChartTitle, nsPrefix)
	}

	// Update axis titles if provided
	if data.CategoryAxisTitle != "" || data.ValueAxisTitle != "" {
		content = updateAxisTitles(content, data.CategoryAxisTitle, data.ValueAxisTitle, nsPrefix)
	}

	// Find chart type and update series
	var err error
	content, err = updateChartSeries(content, data, nsPrefix)
	if err != nil {
		return nil, err
	}

	return []byte(content), nil
}

// detectNamespacePrefix detects the namespace prefix used in the chart XML.
func detectNamespacePrefix(content string) string {
	if strings.Contains(content, "<c:chart") {
		return "c:"
	}
	return ""
}

// updateChartTitle updates the chart title in the XML.
func updateChartTitle(content, title, nsPrefix string) string {
	// Simple approach: find title section and update text
	titleStart := strings.Index(content, "<"+nsPrefix+"title>")
	if titleStart == -1 {
		return content // No title element to update
	}

	titleEnd := strings.Index(content[titleStart:], "</"+nsPrefix+"title>")
	if titleEnd == -1 {
		return content
	}

	// Find the text element within title
	titleSection := content[titleStart : titleStart+titleEnd]
	tStart := strings.Index(titleSection, "<"+nsPrefix+"t>")
	if tStart == -1 {
		return content
	}

	tEnd := strings.Index(titleSection[tStart:], "</"+nsPrefix+"t>")
	if tEnd == -1 {
		return content
	}

	// Replace the title text
	beforeTitle := content[:titleStart+tStart+len("<"+nsPrefix+"t>")]
	afterTitle := content[titleStart+tStart+len("<"+nsPrefix+"t>")+tEnd:]

	return beforeTitle + title + afterTitle
}

// updateAxisTitles updates the axis titles in the chart XML.
func updateAxisTitles(content, catTitle, valTitle, nsPrefix string) string {
	// This is a simplified implementation
	// In production, you'd want more robust XML parsing

	// Update category axis title (usually first axis)
	if catTitle != "" {
		content = updateFirstAxisTitle(content, catTitle, nsPrefix, "catAx")
	}

	// Update value axis title (usually second axis)
	if valTitle != "" {
		content = updateFirstAxisTitle(content, valTitle, nsPrefix, "valAx")
	}

	return content
}

// updateFirstAxisTitle updates the first occurrence of an axis title.
func updateFirstAxisTitle(content, title, nsPrefix, axisType string) string {
	axisStart := strings.Index(content, "<"+nsPrefix+axisType+">")
	if axisStart == -1 {
		return content
	}

	axisEnd := strings.Index(content[axisStart:], "</"+nsPrefix+axisType+">")
	if axisEnd == -1 {
		return content
	}

	axisSection := content[axisStart : axisStart+axisEnd]

	// Find title within this axis
	titleStart := strings.Index(axisSection, "<"+nsPrefix+"title>")
	if titleStart == -1 {
		return content
	}

	tStart := strings.Index(axisSection[titleStart:], "<"+nsPrefix+"t>")
	if tStart == -1 {
		return content
	}

	tEnd := strings.Index(axisSection[titleStart+tStart:], "</"+nsPrefix+"t>")
	if tEnd == -1 {
		return content
	}

	// Replace the axis title text
	absoluteTStart := axisStart + titleStart + tStart + len("<"+nsPrefix+"t>")
	beforeTitle := content[:absoluteTStart]
	afterTitle := content[absoluteTStart+tEnd:]

	return beforeTitle + title + afterTitle
}

// updateChartSeries updates the series data in the chart XML.
func updateChartSeries(content string, data ChartData, nsPrefix string) (string, error) {
	// Find the chart type (barChart, lineChart, scatterChart, pieChart)
	chartTypes := []string{"barChart", "lineChart", "scatterChart", "pieChart", "areaChart"}

	var chartType string
	var chartStart, chartEnd int

	for _, ct := range chartTypes {
		chartStart = strings.Index(content, "<"+nsPrefix+ct+">")
		if chartStart != -1 {
			chartType = ct
			chartEnd = strings.Index(content[chartStart:], "</"+nsPrefix+ct+">")
			if chartEnd == -1 {
				return "", fmt.Errorf("malformed chart XML: no closing tag for %s", ct)
			}
			chartEnd += chartStart
			break
		}
	}

	if chartType == "" {
		return "", fmt.Errorf("unsupported or missing chart type")
	}

	chartSection := content[chartStart:chartEnd]

	// Update or remove series
	updatedSeries, err := updateSeriesSection(chartSection, data, nsPrefix, chartType)
	if err != nil {
		return "", err
	}

	// Reconstruct content
	return content[:chartStart] + updatedSeries + content[chartEnd:], nil
}

// updateSeriesSection updates the series within a chart section.
func updateSeriesSection(chartSection string, data ChartData, nsPrefix, chartType string) (string, error) {
	// Find all series elements
	serTags := findAllSeriesTags(chartSection, nsPrefix)

	if len(serTags) == 0 {
		return "", fmt.Errorf("no series found in chart")
	}
	var scatterXValues []float64
	if chartType == "scatterChart" {
		parsed, err := parseScatterXValues(data.Categories)
		if err != nil {
			return "", err
		}
		scatterXValues = parsed
	}

	// Build new series section
	var buf bytes.Buffer
	buf.WriteString("<" + nsPrefix + chartType + ">")

	// Write series from data
	for i, series := range data.Series {
		serXML, err := buildSeriesXML(series, data.Categories, scatterXValues, i, nsPrefix, chartType)
		if err != nil {
			return "", err
		}
		buf.WriteString(serXML)
	}

	// Copy any non-series elements (like axId, etc.) from the original
	buf.WriteString(copyNonSeriesElements(chartSection, nsPrefix))

	return buf.String(), nil
}

// findAllSeriesTags finds positions of all series tags.
func findAllSeriesTags(content, nsPrefix string) []int {
	var positions []int
	searchStr := "<" + nsPrefix + "ser>"
	offset := 0

	for {
		pos := strings.Index(content[offset:], searchStr)
		if pos == -1 {
			break
		}
		positions = append(positions, offset+pos)
		offset += pos + len(searchStr)
	}

	return positions
}

// buildSeriesXML constructs XML for a single series.
func buildSeriesXML(series SeriesData, categories []string, scatterXValues []float64, idx int, nsPrefix, chartType string) (string, error) {
	var buf bytes.Buffer

	buf.WriteString("<" + nsPrefix + "ser>")
	buf.WriteString("<" + nsPrefix + "idx val=\"" + strconv.Itoa(idx) + "\"/>")
	buf.WriteString("<" + nsPrefix + "order val=\"" + strconv.Itoa(idx) + "\"/>")

	// Series name
	buf.WriteString("<" + nsPrefix + "tx>")
	buf.WriteString("<" + nsPrefix + "v>" + xmlEscape(series.Name) + "</" + nsPrefix + "v>")
	buf.WriteString("</" + nsPrefix + "tx>")

	// Categories (for most chart types except scatter)
	if chartType != "scatterChart" {
		buf.WriteString("<" + nsPrefix + "cat>")
		buf.WriteString("<" + nsPrefix + "strRef>")
		buf.WriteString("<" + nsPrefix + "strCache>")
		buf.WriteString("<" + nsPrefix + "ptCount val=\"" + strconv.Itoa(len(categories)) + "\"/>")
		for i, cat := range categories {
			buf.WriteString("<" + nsPrefix + "pt idx=\"" + strconv.Itoa(i) + "\">")
			buf.WriteString("<" + nsPrefix + "v>" + xmlEscape(cat) + "</" + nsPrefix + "v>")
			buf.WriteString("</" + nsPrefix + "pt>")
		}
		buf.WriteString("</" + nsPrefix + "strCache>")
		buf.WriteString("</" + nsPrefix + "strRef>")
		buf.WriteString("</" + nsPrefix + "cat>")
	} else {
		if len(scatterXValues) == 0 {
			return "", fmt.Errorf("scatter chart categories must be numeric")
		}
		buf.WriteString("<" + nsPrefix + "xVal>")
		buf.WriteString("<" + nsPrefix + "numRef>")
		buf.WriteString("<" + nsPrefix + "numCache>")
		buf.WriteString("<" + nsPrefix + "ptCount val=\"" + strconv.Itoa(len(scatterXValues)) + "\"/>")
		for i, val := range scatterXValues {
			buf.WriteString("<" + nsPrefix + "pt idx=\"" + strconv.Itoa(i) + "\">")
			buf.WriteString("<" + nsPrefix + "v>" + formatFloat(val) + "</" + nsPrefix + "v>")
			buf.WriteString("</" + nsPrefix + "pt>")
		}
		buf.WriteString("</" + nsPrefix + "numCache>")
		buf.WriteString("</" + nsPrefix + "numRef>")
		buf.WriteString("</" + nsPrefix + "xVal>")
	}

	// Values
	valTag := "val"
	if chartType == "scatterChart" {
		valTag = "yVal"
	}

	buf.WriteString("<" + nsPrefix + valTag + ">")
	buf.WriteString("<" + nsPrefix + "numRef>")
	buf.WriteString("<" + nsPrefix + "numCache>")
	buf.WriteString("<" + nsPrefix + "ptCount val=\"" + strconv.Itoa(len(series.Values)) + "\"/>")
	for i, val := range series.Values {
		buf.WriteString("<" + nsPrefix + "pt idx=\"" + strconv.Itoa(i) + "\">")
		buf.WriteString("<" + nsPrefix + "v>" + formatFloat(val) + "</" + nsPrefix + "v>")
		buf.WriteString("</" + nsPrefix + "pt>")
	}
	buf.WriteString("</" + nsPrefix + "numCache>")
	buf.WriteString("</" + nsPrefix + "numRef>")
	buf.WriteString("</" + nsPrefix + valTag + ">")

	buf.WriteString("</" + nsPrefix + "ser>")
	return buf.String(), nil
}

// copyNonSeriesElements copies elements like axId from the original chart.
func copyNonSeriesElements(chartSection, nsPrefix string) string {
	var buf bytes.Buffer

	// Find and copy axId elements
	axIdTags := []string{"axId", "dLbls", "gapWidth", "overlap", "varyColors"}

	for _, tag := range axIdTags {
		startTag := "<" + nsPrefix + tag
		offset := 0

		for {
			pos := strings.Index(chartSection[offset:], startTag)
			if pos == -1 {
				break
			}

			pos += offset
			endPos := strings.Index(chartSection[pos:], "/>")
			if endPos != -1 {
				buf.WriteString(chartSection[pos : pos+endPos+2])
				offset = pos + endPos + 2
			} else {
				endTagStr := "</" + nsPrefix + tag + ">"
				endPos = strings.Index(chartSection[pos:], endTagStr)
				if endPos != -1 {
					buf.WriteString(chartSection[pos : pos+endPos+len(endTagStr)])
					offset = pos + endPos + len(endTagStr)
				} else {
					break
				}
			}
		}
	}

	return buf.String()
}

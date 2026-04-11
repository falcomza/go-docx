package godocx

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// GetChartData reads the categories, series names, and values from chart N (1-based).
func (u *Updater) GetChartData(chartIndex int) (ChartData, error) {
	if u == nil {
		return ChartData{}, fmt.Errorf("updater is nil")
	}
	if chartIndex < 1 {
		return ChartData{}, fmt.Errorf("chart index must be >= 1")
	}
	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	raw, err := os.ReadFile(chartPath)
	if err != nil {
		return ChartData{}, fmt.Errorf("read chart%d.xml: %w", chartIndex, err)
	}
	return parseChartDataFromXML(raw)
}

func parseChartDataFromXML(raw []byte) (ChartData, error) {
	content := string(raw)
	// Basic sanity check — a chart XML file must contain a chartSpace element
	if !strings.Contains(content, "chartSpace") {
		return ChartData{}, fmt.Errorf("content does not appear to be chart XML (missing chartSpace element)")
	}
	// detectNamespacePrefix returns "c:" or "" (already includes the colon)
	ns := detectNamespacePrefix(content)
	// tag builds a tag name like "c:title" or "title"
	tag := func(t string) string {
		return ns + t
	}

	data := ChartData{}

	// Compile the value tag regexp once for the entire parse; the pattern
	// depends on the namespace prefix which is a runtime value.
	vRe := regexp.MustCompile(`<` + regexp.QuoteMeta(tag("v")) + `(?:\s[^>]*)?>([^<]*)<`)

	// Title: first <c:v> inside the <c:title> block
	titleBlock := extractFirstBlock(content, "<"+tag("title")+">", "<"+tag("title")+" ", "</"+tag("title")+">")
	if titleBlock != "" {
		if m := vRe.FindStringSubmatch(titleBlock); m != nil {
			data.ChartTitle = strings.TrimSpace(m[1])
		}
	}

	// Series blocks
	serBlocks := extractBlocks(content, "<"+tag("ser")+">", "<"+tag("ser")+" ", "</"+tag("ser")+">")
	data.Categories = extractCategoriesFromSer(serBlocks, tag, vRe)
	for _, block := range serBlocks {
		name := extractSeriesName(block, tag, vRe)
		values := extractSeriesValues(block, tag, len(data.Categories), vRe)
		data.Series = append(data.Series, SeriesData{Name: name, Values: values})
	}

	return data, nil
}

// extractFirstBlock returns the first block between open and close tags.
func extractFirstBlock(content, openTagExact, openTagAttr, closeTag string) string {
	blocks := extractBlocks(content, openTagExact, openTagAttr, closeTag)
	if len(blocks) == 0 {
		return ""
	}
	return blocks[0]
}

// extractBlocks returns all substrings delimited by open..close tags.
// Tries both exact open tag and attribute-having open tag (e.g. "<w:p>" and "<w:p ").
func extractBlocks(content, openTagExact, openTagAttr, closeTag string) []string {
	var blocks []string
	remaining := content
	for {
		idxExact := strings.Index(remaining, openTagExact)
		idxAttr := strings.Index(remaining, openTagAttr)
		idx := -1
		switch {
		case idxExact >= 0 && idxAttr >= 0:
			if idxExact <= idxAttr {
				idx = idxExact
			} else {
				idx = idxAttr
			}
		case idxExact >= 0:
			idx = idxExact
		case idxAttr >= 0:
			idx = idxAttr
		default:
			return blocks
		}
		closeIdx := strings.Index(remaining[idx:], closeTag)
		if closeIdx < 0 {
			return blocks
		}
		end := idx + closeIdx + len(closeTag)
		blocks = append(blocks, remaining[idx:end])
		remaining = remaining[end:]
	}
}

func extractSeriesName(block string, tag func(string) string, vRe *regexp.Regexp) string {
	txClose := "</" + tag("tx") + ">"
	txEnd := strings.Index(block, txClose)
	if txEnd < 0 {
		return ""
	}
	txBlock := block[:txEnd]
	if m := vRe.FindStringSubmatch(txBlock); m != nil {
		return strings.TrimSpace(m[1])
	}
	return ""
}

func extractCategoriesFromSer(serBlocks []string, tag func(string) string, vRe *regexp.Regexp) []string {
	if len(serBlocks) == 0 {
		return nil
	}
	first := serBlocks[0]

	// Standard charts use <cat>; scatter charts typically use <xVal>.
	if cats := extractStringValuesFromDataRef(first, tag, "cat", vRe); len(cats) > 0 {
		return cats
	}
	if xVals := extractStringValuesFromDataRef(first, tag, "xVal", vRe); len(xVals) > 0 {
		return xVals
	}

	return nil
}

func extractSeriesValues(block string, tag func(string) string, count int, vRe *regexp.Regexp) []float64 {
	// Standard charts use <val>; scatter charts typically use <yVal>.
	values := extractNumericValuesFromDataRef(block, tag, "val", vRe)
	if len(values) == 0 {
		values = extractNumericValuesFromDataRef(block, tag, "yVal", vRe)
	}
	for len(values) < count {
		values = append(values, 0)
	}
	if len(values) > count && count > 0 {
		values = values[:count]
	}
	return values
}

func extractStringValuesFromDataRef(block string, tag func(string) string, refTag string, vRe *regexp.Regexp) []string {
	open := "<" + tag(refTag) + ">"
	close := "</" + tag(refTag) + ">"
	start := strings.Index(block, open)
	end := strings.LastIndex(block, close)
	if start < 0 || end < 0 {
		return nil
	}
	dataBlock := block[start : end+len(close)]
	matches := vRe.FindAllStringSubmatch(dataBlock, -1)
	vals := make([]string, 0, len(matches))
	for _, m := range matches {
		vals = append(vals, strings.TrimSpace(m[1]))
	}
	return vals
}

func extractNumericValuesFromDataRef(block string, tag func(string) string, refTag string, vRe *regexp.Regexp) []float64 {
	stringVals := extractStringValuesFromDataRef(block, tag, refTag, vRe)
	values := make([]float64, 0, len(stringVals))
	for _, v := range stringVals {
		f, err := strconv.ParseFloat(v, 64)
		if err != nil {
			f = 0
		}
		values = append(values, f)
	}
	return values
}

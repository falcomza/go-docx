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
	catOpen := "<" + tag("cat") + ">"
	catClose := "</" + tag("cat") + ">"
	catStart := strings.Index(serBlocks[0], catOpen)
	catEnd := strings.LastIndex(serBlocks[0], catClose)
	if catStart < 0 || catEnd < 0 {
		return nil
	}
	catBlock := serBlocks[0][catStart : catEnd+len(catClose)]
	matches := vRe.FindAllStringSubmatch(catBlock, -1)
	cats := make([]string, 0, len(matches))
	for _, m := range matches {
		cats = append(cats, strings.TrimSpace(m[1]))
	}
	return cats
}

func extractSeriesValues(block string, tag func(string) string, count int, vRe *regexp.Regexp) []float64 {
	valOpen := "<" + tag("val") + ">"
	valClose := "</" + tag("val") + ">"
	valStart := strings.Index(block, valOpen)
	valEnd := strings.LastIndex(block, valClose)
	if valStart < 0 || valEnd < 0 {
		return make([]float64, count)
	}
	valBlock := block[valStart : valEnd+len(valClose)]
	matches := vRe.FindAllStringSubmatch(valBlock, -1)
	values := make([]float64, 0, len(matches))
	for _, m := range matches {
		f, _ := strconv.ParseFloat(strings.TrimSpace(m[1]), 64)
		values = append(values, f)
	}
	for len(values) < count {
		values = append(values, 0)
	}
	if len(values) > count && count > 0 {
		values = values[:count]
	}
	return values
}

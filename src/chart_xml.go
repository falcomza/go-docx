package docxchartupdater

import (
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"
)

var (
	reSeriesBlock = regexp.MustCompile(`(?s)<c:ser\b.*?</c:ser>`)
	reTxV         = regexp.MustCompile(`(?s)(<c:tx\b.*?<c:v>)(.*?)(</c:v>)`)
	reCatCache    = regexp.MustCompile(`(?s)(<c:cat\b.*?<c:strCache>)(.*?)(</c:strCache>.*?</c:cat>)`)
	reValCache    = regexp.MustCompile(`(?s)(<c:val\b.*?<c:numCache>)(.*?)(</c:numCache>.*?</c:val>)`)
)

func updateChartXML(chartPath string, data ChartData) error {
	raw, err := os.ReadFile(chartPath)
	if err != nil {
		return fmt.Errorf("read chart xml: %w", err)
	}

	content := string(raw)
	seriesBlocks := reSeriesBlock.FindAllStringIndex(content, -1)
	if len(seriesBlocks) == 0 {
		return fmt.Errorf("no <c:ser> blocks found in chart xml")
	}

	var out strings.Builder
	prev := 0

	for i, bounds := range seriesBlocks {
		out.WriteString(content[prev:bounds[0]])
		block := content[bounds[0]:bounds[1]]

		if i < len(data.Series) {
			block = updateSeriesBlock(block, data.Categories, data.Series[i])
		}

		out.WriteString(block)
		prev = bounds[1]
	}
	out.WriteString(content[prev:])

	if err := os.WriteFile(chartPath, []byte(out.String()), 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

func updateSeriesBlock(block string, categories []string, series SeriesData) string {
	xmlEscapedName := xmlEscape(series.Name)
	block = reTxV.ReplaceAllString(block, "$1"+xmlEscapedName+"$3")

	block = reCatCache.ReplaceAllStringFunc(block, func(m string) string {
		parts := reCatCache.FindStringSubmatch(m)
		if len(parts) != 4 {
			return m
		}
		return parts[1] + buildStrCache(categories) + parts[3]
	})

	block = reValCache.ReplaceAllStringFunc(block, func(m string) string {
		parts := reValCache.FindStringSubmatch(m)
		if len(parts) != 4 {
			return m
		}
		return parts[1] + buildNumCache(series.Values) + parts[3]
	})

	return block
}

func buildStrCache(values []string) string {
	var b strings.Builder
	b.WriteString(`<c:ptCount val="` + strconv.Itoa(len(values)) + `"/>`)
	for i, v := range values {
		b.WriteString(`<c:pt idx="` + strconv.Itoa(i) + `"><c:v>` + xmlEscape(v) + `</c:v></c:pt>`)
	}
	return b.String()
}

func buildNumCache(values []float64) string {
	var b strings.Builder
	b.WriteString(`<c:ptCount val="` + strconv.Itoa(len(values)) + `"/>`)
	for i, v := range values {
		b.WriteString(`<c:pt idx="` + strconv.Itoa(i) + `"><c:v>` + strconv.FormatFloat(v, 'f', -1, 64) + `</c:v></c:pt>`)
	}
	return b.String()
}

func xmlEscape(s string) string {
	replacer := strings.NewReplacer(
		"&", "&amp;",
		"<", "&lt;",
		">", "&gt;",
		`"`, "&quot;",
		"'", "&apos;",
	)
	return replacer.Replace(s)
}

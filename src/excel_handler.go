package docxchartupdater

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

func updateEmbeddedWorkbook(xlsxPath string, data ChartData) error {
	xlsxRaw, err := os.ReadFile(xlsxPath)
	if err != nil {
		return fmt.Errorf("read embedded workbook: %w", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(xlsxRaw), int64(len(xlsxRaw)))
	if err != nil {
		return fmt.Errorf("open workbook zip: %w", err)
	}

	entries := make(map[string][]byte, len(zr.File))
	names := make([]string, 0, len(zr.File))

	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return fmt.Errorf("open workbook entry %s: %w", f.Name, err)
		}
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return fmt.Errorf("read workbook entry %s: %w", f.Name, err)
		}
		entries[f.Name] = content
		names = append(names, f.Name)
	}

	sort.Strings(names)
	worksheetPath, err := resolveWorksheetPath(entries, names)
	if err != nil {
		return err
	}
	if worksheetPath == "" {
		return fmt.Errorf("no worksheet found in embedded workbook")
	}

	useSharedStrings := false
	stringIndexes := map[string]int{}
	if sharedStringsRaw, ok := entries["xl/sharedStrings.xml"]; ok {
		useSharedStrings = true
		updatedSharedStrings, indexes, err := updateSharedStringsXML(sharedStringsRaw, data)
		if err != nil {
			return err
		}
		entries["xl/sharedStrings.xml"] = updatedSharedStrings
		stringIndexes = indexes
	}

	updatedWorksheet, err := updateWorksheetXML(entries[worksheetPath], data, useSharedStrings, stringIndexes)
	if err != nil {
		return err
	}
	entries[worksheetPath] = updatedWorksheet

	buf := &bytes.Buffer{}
	zw := zip.NewWriter(buf)
	for _, name := range names {
		w, err := zw.Create(name)
		if err != nil {
			return fmt.Errorf("create workbook entry %s: %w", name, err)
		}
		if _, err := w.Write(entries[name]); err != nil {
			return fmt.Errorf("write workbook entry %s: %w", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		return fmt.Errorf("close workbook writer: %w", err)
	}

	if err := os.MkdirAll(filepath.Dir(xlsxPath), 0o755); err != nil {
		return fmt.Errorf("create workbook parent dir: %w", err)
	}
	if err := os.WriteFile(xlsxPath, buf.Bytes(), 0o644); err != nil {
		return fmt.Errorf("write embedded workbook: %w", err)
	}

	return nil
}

func firstWorksheetPath(names []string) string {
	for _, name := range names {
		if strings.HasPrefix(name, "xl/worksheets/sheet") && strings.HasSuffix(name, ".xml") {
			return name
		}
	}
	return ""
}

func resolveWorksheetPath(entries map[string][]byte, names []string) (string, error) {
	const workbookPath = "xl/workbook.xml"
	const relsPath = "xl/_rels/workbook.xml.rels"
	workbookRaw, okWorkbook := entries[workbookPath]
	relsRaw, okRels := entries[relsPath]
	if !okWorkbook || !okRels {
		return firstWorksheetPath(names), nil
	}

	var wb workbookXML
	if err := xml.Unmarshal(workbookRaw, &wb); err != nil {
		return "", fmt.Errorf("parse workbook.xml: %w", err)
	}
	if len(wb.Sheets) == 0 {
		return firstWorksheetPath(names), nil
	}

	var rels relationships
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return "", fmt.Errorf("parse workbook.xml.rels: %w", err)
	}

	target := ""
	for _, rel := range rels.Relationships {
		if rel.ID == wb.Sheets[0].RelID {
			target = rel.Target
			break
		}
	}
	if target == "" {
		return firstWorksheetPath(names), nil
	}

	full := filepath.ToSlash(filepath.Clean(filepath.Join("xl", target)))
	if _, ok := entries[full]; ok {
		return full, nil
	}
	return firstWorksheetPath(names), nil
}

type workbookXML struct {
	Sheets []workbookSheet `xml:"sheets>sheet"`
}

type workbookSheet struct {
	RelID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

func updateWorksheetXML(existing []byte, data ChartData, useSharedStrings bool, stringIndexes map[string]int) ([]byte, error) {
	updated := string(existing)
	newSheetData := buildSheetDataXML(data, useSharedStrings, stringIndexes)

	reSheetData := regexp.MustCompile(`(?s)<sheetData\b[^>]*>.*?</sheetData>`)
	if !reSheetData.MatchString(updated) {
		return nil, fmt.Errorf("worksheet has no sheetData element")
	}
	updated = reSheetData.ReplaceAllString(updated, newSheetData)

	lastCol := columnLetters(len(data.Series) + 1)
	lastRow := len(data.Categories) + 1
	newDimension := `<dimension ref="A1:` + lastCol + strconv.Itoa(lastRow) + `"/>`

	reDimension := regexp.MustCompile(`<dimension\b[^>]*ref=\"[^\"]*\"[^>]*/>`)
	if reDimension.MatchString(updated) {
		updated = reDimension.ReplaceAllString(updated, newDimension)
	}

	return []byte(updated), nil
}

func buildSheetDataXML(data ChartData, useSharedStrings bool, stringIndexes map[string]int) string {
	var b strings.Builder
	b.WriteString(`<sheetData>`)

	// Header row: A1 blank, then series names.
	b.WriteString(`<row r="1">`)
	b.WriteString(stringCell("A1", "", useSharedStrings, stringIndexes))
	for i, s := range data.Series {
		b.WriteString(stringCell(cellRef(i+2, 1), s.Name, useSharedStrings, stringIndexes))
	}
	b.WriteString(`</row>`)

	for rowIdx, cat := range data.Categories {
		r := rowIdx + 2
		b.WriteString(`<row r="` + strconv.Itoa(r) + `">`)
		b.WriteString(stringCell(cellRef(1, r), cat, useSharedStrings, stringIndexes))
		for sIdx, s := range data.Series {
			value := 0.0
			if rowIdx < len(s.Values) {
				value = s.Values[rowIdx]
			}
			b.WriteString(numberCell(cellRef(sIdx+2, r), value))
		}
		b.WriteString(`</row>`)
	}

	b.WriteString(`</sheetData>`)
	return b.String()
}

func stringCell(ref, value string, useSharedStrings bool, stringIndexes map[string]int) string {
	if useSharedStrings {
		if idx, ok := stringIndexes[value]; ok {
			return sharedStringCell(ref, idx)
		}
	}
	return inlineStringCell(ref, value)
}

func inlineStringCell(ref, value string) string {
	return `<c r="` + ref + `" t="inlineStr"><is><t>` + xmlEscape(value) + `</t></is></c>`
}

func sharedStringCell(ref string, stringIndex int) string {
	return `<c r="` + ref + `" t="s"><v>` + strconv.Itoa(stringIndex) + `</v></c>`
}

func numberCell(ref string, value float64) string {
	return `<c r="` + ref + `"><v>` + strconv.FormatFloat(value, 'f', -1, 64) + `</v></c>`
}

func cellRef(col, row int) string {
	return columnLetters(col) + strconv.Itoa(row)
}

func columnLetters(n int) string {
	if n <= 0 {
		return "A"
	}
	var out []byte
	for n > 0 {
		n--
		out = append([]byte{byte('A' + (n % 26))}, out...)
		n /= 26
	}
	return string(out)
}

type sharedStringTable struct {
	XMLName     xml.Name           `xml:"sst"`
	XMLNS       string             `xml:"xmlns,attr,omitempty"`
	Count       int                `xml:"count,attr"`
	UniqueCount int                `xml:"uniqueCount,attr"`
	SI          []sharedStringItem `xml:"si"`
}

type sharedStringItem struct {
	T string `xml:"t"`
}

func updateSharedStringsXML(existing []byte, data ChartData) ([]byte, map[string]int, error) {
	var parsed sharedStringTable
	if err := xml.Unmarshal(existing, &parsed); err != nil {
		return nil, nil, fmt.Errorf("parse sharedStrings.xml: %w", err)
	}
	if parsed.XMLNS == "" {
		parsed.XMLNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	}

	// Preserve existing string table entries and only append missing values.
	indexes := make(map[string]int, len(parsed.SI)+1+len(data.Series)+len(data.Categories))
	for i, item := range parsed.SI {
		if _, exists := indexes[item.T]; !exists {
			indexes[item.T] = i
		}
	}

	appendIfMissing := func(value string) {
		if _, exists := indexes[value]; exists {
			return
		}
		parsed.SI = append(parsed.SI, sharedStringItem{T: value})
		indexes[value] = len(parsed.SI) - 1
	}

	appendIfMissing("")
	for _, s := range data.Series {
		appendIfMissing(s.Name)
	}
	for _, c := range data.Categories {
		appendIfMissing(c)
	}
	parsed.Count = len(parsed.SI)
	parsed.UniqueCount = len(parsed.SI)

	encoded, err := xml.Marshal(parsed)
	if err != nil {
		return nil, nil, fmt.Errorf("marshal sharedStrings.xml: %w", err)
	}

	out := append([]byte(xml.Header), encoded...)
	return out, indexes, nil
}

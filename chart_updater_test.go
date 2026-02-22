package godocx_test

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

func TestBasicUpdate(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	data := godocx.ChartData{
		Categories: []string{"Device A", "Device B", "Device C"},
		Series: []godocx.SeriesData{
			{Name: "Critical", Values: []float64{4, 3, 2}},
			{Name: "Non-critical", Values: []float64{8, 7, 6}},
		},
	}

	if err := u.UpdateChart(1, data); err != nil {
		t.Fatalf("UpdateChart failed: %v", err)
	}
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	chartXML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	if !strings.Contains(chartXML, `<v>Device A</v>`) && !strings.Contains(chartXML, `<c:v>Device A</c:v>`) {
		t.Fatalf("chart xml missing updated category")
	}
	if !strings.Contains(chartXML, `<v>8</v>`) && !strings.Contains(chartXML, `<c:v>8</c:v>`) {
		t.Fatalf("chart xml missing updated value")
	}

	xlsxRaw := readZipEntryBytes(t, outputPath, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx")
	sheetXML := readWorkbookEntry(t, xlsxRaw, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheetXML, `r="B1"`) || !strings.Contains(sheetXML, `Critical`) {
		t.Fatalf("worksheet missing series header")
	}
	if !strings.Contains(sheetXML, `r="A2"`) || !strings.Contains(sheetXML, `Device A`) {
		t.Fatalf("worksheet missing category data")
	}
	if !strings.Contains(sheetXML, `<c r="C4"><v>6</v></c>`) {
		t.Fatalf("worksheet missing numeric data")
	}
}

func TestInvalidData(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.UpdateChart(1, godocx.ChartData{
		Categories: []string{"A", "B"},
		Series:     []godocx.SeriesData{{Name: "Critical", Values: []float64{1}}},
	})
	if err == nil {
		t.Fatalf("expected length mismatch error")
	}
}

func TestUpdateWithSharedStringsWorkbook(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxWithSharedStrings(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	data := godocx.ChartData{
		Categories: []string{"Node 1", "Node 2"},
		Series: []godocx.SeriesData{
			{Name: "Critical", Values: []float64{11, 9}},
			{Name: "Non-critical", Values: []float64{22, 18}},
		},
	}

	if err := u.UpdateChart(1, data); err != nil {
		t.Fatalf("UpdateChart failed: %v", err)
	}
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	xlsxRaw := readZipEntryBytes(t, outputPath, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx")
	sheetXML := readWorkbookEntry(t, xlsxRaw, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheetXML, `t="s"`) {
		t.Fatalf("worksheet did not use shared string cells")
	}
	if !strings.Contains(sheetXML, `<c r="A2" t="s"><v>`) {
		t.Fatalf("worksheet missing shared string reference")
	}

	sharedStringsXML := readWorkbookEntry(t, xlsxRaw, "xl/sharedStrings.xml")
	if !strings.Contains(sharedStringsXML, "<t>placeholder</t>") {
		t.Fatalf("sharedStrings should preserve existing values")
	}
	if !strings.Contains(sharedStringsXML, "<t>Node 1</t>") {
		t.Fatalf("sharedStrings missing category text")
	}
	if !strings.Contains(sharedStringsXML, "<t>Critical</t>") {
		t.Fatalf("sharedStrings missing series name")
	}
}

func TestUpdateSpecificChartInMultiChartDocx(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input_multi.docx")
	outputPath := filepath.Join(tempDir, "output_multi.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxTwoCharts(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	data := godocx.ChartData{
		Categories: []string{"Router A", "Router B"},
		Series: []godocx.SeriesData{
			{Name: "Critical", Values: []float64{5, 7}},
			{Name: "Non-critical", Values: []float64{15, 17}},
		},
	}

	if err := u.UpdateChart(2, data); err != nil {
		t.Fatalf("UpdateChart(2) failed: %v", err)
	}
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	chart1XML := readZipEntry(t, outputPath, "word/charts/chart1.xml")
	chart2XML := readZipEntry(t, outputPath, "word/charts/chart2.xml")
	if !strings.Contains(chart1XML, "<v>Old 1</v>") && !strings.Contains(chart1XML, "<c:v>Old 1</c:v>") {
		t.Fatalf("chart1 should remain unchanged")
	}
	if !strings.Contains(chart2XML, "<v>Router A</v>") && !strings.Contains(chart2XML, "<c:v>Router A</c:v>") {
		t.Fatalf("chart2 should contain updated data")
	}

	workbook1Raw := readZipEntryBytes(t, outputPath, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx")
	workbook2Raw := readZipEntryBytes(t, outputPath, "word/embeddings/Microsoft_Excel_Worksheet2.xlsx")

	workbook1Sheet := readWorkbookEntry(t, workbook1Raw, "xl/worksheets/sheet1.xml")
	workbook2Sheet := readWorkbookEntry(t, workbook2Raw, "xl/worksheets/sheet1.xml")
	if strings.Contains(workbook1Sheet, "Router A") {
		t.Fatalf("workbook 1 should remain unchanged")
	}
	if !strings.Contains(workbook2Sheet, "Router A") {
		t.Fatalf("workbook 2 should be updated")
	}
}

func buildFixtureDocx(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
	addZipEntry(t, docxZip, "word/charts/chart1.xml", chartFixtureXML)
	addZipEntry(t, docxZip, "word/charts/_rels/chart1.xml.rels", chartRelsFixtureXML)
	addZipEntryBytes(t, docxZip, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx", buildFixtureWorkbook(t))

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func buildFixtureDocxWithSharedStrings(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
	addZipEntry(t, docxZip, "word/charts/chart1.xml", chartFixtureXML)
	addZipEntry(t, docxZip, "word/charts/_rels/chart1.xml.rels", chartRelsFixtureXML)
	addZipEntryBytes(t, docxZip, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx", buildFixtureWorkbookWithSharedStrings(t))

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func buildFixtureDocxTwoCharts(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)
	addZipEntry(t, docxZip, "word/charts/chart1.xml", chartFixtureXML)
	addZipEntry(t, docxZip, "word/charts/chart2.xml", chart2FixtureXML)
	addZipEntry(t, docxZip, "word/charts/_rels/chart1.xml.rels", chartRelsFixtureXML)
	addZipEntry(t, docxZip, "word/charts/_rels/chart2.xml.rels", chart2RelsFixtureXML)
	addZipEntryBytes(t, docxZip, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx", buildFixtureWorkbook(t))
	addZipEntryBytes(t, docxZip, "word/embeddings/Microsoft_Excel_Worksheet2.xlsx", buildFixtureWorkbook(t))

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func buildFixtureDocxNoCharts(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func buildFixtureWorkbook(t *testing.T) []byte {
	t.Helper()

	buf := &bytes.Buffer{}
	w := zip.NewWriter(buf)

	addZipEntry(t, w, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, w, "xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>`)
	addZipEntry(t, w, "xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZipEntry(t, w, "xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`)

	if err := w.Close(); err != nil {
		t.Fatalf("close workbook zip: %v", err)
	}
	return buf.Bytes()
}

func buildFixtureWorkbookWithSharedStrings(t *testing.T) []byte {
	t.Helper()

	buf := &bytes.Buffer{}
	w := zip.NewWriter(buf)

	addZipEntry(t, w, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, w, "xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>`)
	addZipEntry(t, w, "xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZipEntry(t, w, "xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`)
	addZipEntry(t, w, "xl/sharedStrings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>placeholder</t></si></sst>`)

	if err := w.Close(); err != nil {
		t.Fatalf("close workbook zip: %v", err)
	}
	return buf.Bytes()
}

func addZipEntry(t *testing.T, w *zip.Writer, path, content string) {
	t.Helper()
	entry, err := w.Create(path)
	if err != nil {
		t.Fatalf("create zip entry %s: %v", path, err)
	}
	if _, err := entry.Write([]byte(content)); err != nil {
		t.Fatalf("write zip entry %s: %v", path, err)
	}
}

func addZipEntryBytes(t *testing.T, w *zip.Writer, path string, content []byte) {
	t.Helper()
	entry, err := w.Create(path)
	if err != nil {
		t.Fatalf("create zip entry %s: %v", path, err)
	}
	if _, err := entry.Write(content); err != nil {
		t.Fatalf("write zip entry %s: %v", path, err)
	}
}

func readZipEntry(t *testing.T, zipPath, entryPath string) string {
	t.Helper()
	return string(readZipEntryBytes(t, zipPath, entryPath))
}

func readZipEntryBytes(t *testing.T, zipPath, entryPath string) []byte {
	t.Helper()

	r, err := zip.OpenReader(zipPath)
	if err != nil {
		t.Fatalf("open zip %s: %v", zipPath, err)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == entryPath {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("open entry %s: %v", entryPath, err)
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("read entry %s: %v", entryPath, err)
			}
			return b
		}
	}

	t.Fatalf("entry not found: %s", entryPath)
	return nil
}

func readWorkbookEntry(t *testing.T, workbookRaw []byte, entryPath string) string {
	t.Helper()

	r, err := zip.NewReader(bytes.NewReader(workbookRaw), int64(len(workbookRaw)))
	if err != nil {
		t.Fatalf("open workbook zip: %v", err)
	}
	for _, f := range r.File {
		if f.Name != entryPath {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open workbook entry %s: %v", entryPath, err)
		}
		defer rc.Close()
		b, err := io.ReadAll(rc)
		if err != nil {
			t.Fatalf("read workbook entry %s: %v", entryPath, err)
		}
		return string(b)
	}
	t.Fatalf("workbook entry not found: %s", entryPath)
	return ""
}

func TestGetChartCount(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	count, err := u.GetChartCount()
	if err != nil {
		t.Fatalf("GetChartCount: %v", err)
	}
	if count != 1 {
		t.Errorf("expected 1 chart, got %d", count)
	}
}

func TestGetChartCountTwoCharts(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	if err := os.WriteFile(inputPath, buildFixtureDocxTwoCharts(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	count, err := u.GetChartCount()
	if err != nil {
		t.Fatalf("GetChartCount: %v", err)
	}
	if count != 2 {
		t.Errorf("expected 2 charts, got %d", count)
	}
}

func TestGetChartCountEmpty(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	if err := os.WriteFile(inputPath, buildFixtureDocxNoCharts(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	count, err := u.GetChartCount()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if count != 0 {
		t.Errorf("expected 0, got %d", count)
	}
}

func TestGetChartCountWithTemplate(t *testing.T) {
	templatePath := filepath.Join("templates", "docx_template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skip("Template file not found")
	}

	u, err := godocx.New(templatePath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	count, err := u.GetChartCount()
	if err != nil {
		t.Fatalf("GetChartCount: %v", err)
	}
	if count < 1 {
		t.Errorf("expected at least 1 chart, got %d", count)
	}
}

const chartRelsFixtureXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/>
</Relationships>`

const chart2RelsFixtureXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet2.xlsx"/>
</Relationships>`

const chartFixtureXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx><c:v>Critical</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Old 1</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:tx><c:v>Non-critical</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Old 1</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
  <c:externalData r:id="rId1"/>
</c:chartSpace>`

const chart2FixtureXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx><c:v>Critical</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Old X</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:tx><c:v>Non-critical</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Old X</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
  <c:externalData r:id="rId1"/>
</c:chartSpace>`

func TestGetChartData(t *testing.T) {
	templatePath := filepath.Join("templates", "docx_template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skip("Template file not found")
	}

	u, err := godocx.New(templatePath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	count, err := u.GetChartCount()
	if err != nil {
		t.Fatalf("GetChartCount: %v", err)
	}
	if count == 0 {
		t.Skip("Template has no charts")
	}

	data, err := u.GetChartData(1)
	if err != nil {
		t.Fatalf("GetChartData: %v", err)
	}
	if len(data.Categories) == 0 {
		t.Error("expected at least one category")
	}
	if len(data.Series) == 0 {
		t.Error("expected at least one series")
	}
	for i, s := range data.Series {
		if s.Name == "" {
			t.Errorf("series[%d] has empty name", i)
		}
		if len(s.Values) != len(data.Categories) {
			t.Errorf("series[%d]: %d values for %d categories", i, len(s.Values), len(data.Categories))
		}
	}
}

func TestGetChartDataInvalidIndex(t *testing.T) {
	templatePath := filepath.Join("templates", "docx_template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skip("Template file not found")
	}

	u, err := godocx.New(templatePath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	_, err = u.GetChartData(0)
	if err == nil {
		t.Error("expected error for index 0")
	}
	_, err = u.GetChartData(999)
	if err == nil {
		t.Error("expected error for non-existent chart")
	}
}

func TestGetChartDataRoundTrip(t *testing.T) {
	// InsertChart then GetChartData should return the same categories and series names
	templatePath := filepath.Join("templates", "docx_template.docx")
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skip("Template file not found")
	}

	u, err := godocx.New(templatePath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	wantCategories := []string{"Jan", "Feb", "Mar"}
	wantSeries := []godocx.SeriesOptions{
		{Name: "Alpha", Values: []float64{10, 20, 30}},
		{Name: "Beta", Values: []float64{40, 50, 60}},
	}

	err = u.InsertChart(godocx.ChartOptions{
		Categories: wantCategories,
		Series:     wantSeries,
		Position:   godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertChart: %v", err)
	}

	count, _ := u.GetChartCount()
	data, err := u.GetChartData(count) // newest chart is last
	if err != nil {
		t.Fatalf("GetChartData: %v", err)
	}

	if len(data.Categories) != len(wantCategories) {
		t.Errorf("categories: got %d want %d", len(data.Categories), len(wantCategories))
	}
	for i, cat := range wantCategories {
		if i >= len(data.Categories) {
			break
		}
		if data.Categories[i] != cat {
			t.Errorf("category[%d]: got %q want %q", i, data.Categories[i], cat)
		}
	}
	if len(data.Series) != len(wantSeries) {
		t.Errorf("series count: got %d want %d", len(data.Series), len(wantSeries))
	}
	for i, s := range wantSeries {
		if i >= len(data.Series) {
			break
		}
		if data.Series[i].Name != s.Name {
			t.Errorf("series[%d].Name: got %q want %q", i, data.Series[i].Name, s.Name)
		}
	}
}

func TestGetChartDataFromFixture(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	data, err := u.GetChartData(1)
	if err != nil {
		t.Fatalf("GetChartData: %v", err)
	}

	if len(data.Categories) != 1 {
		t.Errorf("expected 1 category, got %d", len(data.Categories))
	}
	if len(data.Categories) > 0 && data.Categories[0] != "Old 1" {
		t.Errorf("category[0]: got %q want %q", data.Categories[0], "Old 1")
	}
	if len(data.Series) != 2 {
		t.Errorf("expected 2 series, got %d", len(data.Series))
	}
	if len(data.Series) > 0 && data.Series[0].Name != "Critical" {
		t.Errorf("series[0].Name: got %q want %q", data.Series[0].Name, "Critical")
	}
	if len(data.Series) > 1 && data.Series[1].Name != "Non-critical" {
		t.Errorf("series[1].Name: got %q want %q", data.Series[1].Name, "Non-critical")
	}
	if len(data.Series) > 0 && len(data.Series[0].Values) > 0 && data.Series[0].Values[0] != 1 {
		t.Errorf("series[0].Values[0]: got %v want 1", data.Series[0].Values[0])
	}
	if len(data.Series) > 1 && len(data.Series[1].Values) > 0 && data.Series[1].Values[0] != 2 {
		t.Errorf("series[1].Values[0]: got %v want 2", data.Series[1].Values[0])
	}
}

func TestGetChartDataInvalidIndexFromFixture(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	_, err = u.GetChartData(0)
	if err == nil {
		t.Error("expected error for index 0")
	}

	_, err = u.GetChartData(999)
	if err == nil {
		t.Error("expected error for non-existent chart index 999")
	}
}

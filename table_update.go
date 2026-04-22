package godocx

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// UpdateTableCell replaces the text content of a cell in an existing table.
// tableIndex, row, and col are all 1-based.
//
// Note: nested tables (a table inside a table cell) are not supported.
func (u *Updater) UpdateTableCell(tableIndex, row, col int, value string) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if tableIndex < 1 {
		return fmt.Errorf("tableIndex must be >= 1")
	}
	if row < 1 {
		return fmt.Errorf("row must be >= 1")
	}
	if col < 1 {
		return fmt.Errorf("col must be >= 1")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := updateTableCellContent(raw, tableIndex, row, col, value)
	if err != nil {
		return err
	}

	return os.WriteFile(docPath, updated, 0o644)
}

// updateTableCellContent performs the XML surgery.
func updateTableCellContent(raw []byte, tableIndex, row, col int, value string) ([]byte, error) {
	content := string(raw)

	// Locate Nth <w:tbl> block
	tblStart, tblEnd, err := findNthXMLBlock(content, "w:tbl", tableIndex)
	if err != nil {
		return nil, fmt.Errorf("table %d not found: %w", tableIndex, err)
	}
	tblContent := content[tblStart:tblEnd]

	// Locate Rth <w:tr> within table
	trStart, trEnd, err := findNthXMLBlock(tblContent, "w:tr", row)
	if err != nil {
		return nil, fmt.Errorf("table %d row %d not found: %w", tableIndex, row, err)
	}
	trContent := tblContent[trStart:trEnd]

	// Locate Cth <w:tc> within row
	tcStart, tcEnd, err := findNthXMLBlock(trContent, "w:tc", col)
	if err != nil {
		return nil, fmt.Errorf("table %d row %d col %d not found: %w", tableIndex, row, col, err)
	}
	tcContent := trContent[tcStart:tcEnd]

	// Replace cell text
	updatedTC, err := replaceCellText(tcContent, value)
	if err != nil {
		return nil, err
	}

	// Reassemble
	newTR := trContent[:tcStart] + updatedTC + trContent[tcEnd:]
	newTbl := tblContent[:trStart] + newTR + tblContent[trEnd:]
	result := content[:tblStart] + newTbl + content[tblEnd:]
	return []byte(result), nil
}

// findNthXMLBlock finds the start and end offset of the Nth occurrence of <tag>...</tag>
// within content. Handles both <tag> and <tag attr="..."> open-tag forms.
//
// Note: does not handle nested same-tag elements (e.g. a table inside a table cell).
// For UpdateTableCell, this means cells in nested tables cannot be addressed by index.
func findNthXMLBlock(content, tag string, n int) (start, end int, err error) {
	openExact := "<" + tag + ">"
	openAttr := "<" + tag + " "
	closeTag := "</" + tag + ">"

	count := 0
	remaining := content
	offset := 0
	for {
		ie := strings.Index(remaining, openExact)
		ia := strings.Index(remaining, openAttr)
		idx := -1
		switch {
		case ie >= 0 && ia >= 0:
			if ie <= ia {
				idx = ie
			} else {
				idx = ia
			}
		case ie >= 0:
			idx = ie
		case ia >= 0:
			idx = ia
		}
		if idx < 0 {
			return 0, 0, fmt.Errorf("only %d %s element(s) found", count, tag)
		}
		count++
		absStart := offset + idx
		closeIdx := strings.Index(remaining[idx:], closeTag)
		if closeIdx < 0 {
			return 0, 0, fmt.Errorf("unclosed <%s>", tag)
		}
		absEnd := absStart + closeIdx + len(closeTag)
		if count == n {
			return absStart, absEnd, nil
		}
		advance := idx + closeIdx + len(closeTag)
		offset += advance
		remaining = remaining[advance:]
	}
}

// replaceCellText replaces all run text inside a <w:tc> with value, preserving <w:tcPr>.
func replaceCellText(tcContent, value string) (string, error) {
	escaped := xmlEscape(value)

	// Preserve <w:tcPr> cell properties block if present
	tcPr := ""
	if start := strings.Index(tcContent, "<w:tcPr>"); start >= 0 {
		if end := strings.Index(tcContent[start:], "</w:tcPr>"); end >= 0 {
			tcPr = tcContent[start : start+end+len("</w:tcPr>")]
		}
	} else if start := strings.Index(tcContent, "<w:tcPr "); start >= 0 {
		if end := strings.Index(tcContent[start:], "</w:tcPr>"); end >= 0 {
			tcPr = tcContent[start : start+end+len("</w:tcPr>")]
		}
	}

	// Preserve <w:pPr> paragraph properties if present in the first <w:p>
	pPr := ""
	pStart, pEnd, err := findNthXMLBlock(tcContent, "w:p", 1)
	if err != nil {
		// No paragraph at all — create a minimal one
		openTag := "<w:tc>"
		if strings.Contains(tcContent, "<w:tc ") {
			// Use the actual opening tag
			end := strings.Index(tcContent, ">")
			if end >= 0 {
				openTag = tcContent[:end+1]
			}
		}
		var b strings.Builder
		b.WriteString(openTag)
		b.WriteString(tcPr)
		if escaped != "" {
			b.WriteString(`<w:p><w:r><w:t xml:space="preserve">`)
			b.WriteString(escaped)
			b.WriteString(`</w:t></w:r></w:p>`)
		} else {
			b.WriteString("<w:p/>")
		}
		b.WriteString("</w:tc>")
		return b.String(), nil
	}
	pContent := tcContent[pStart:pEnd]
	if ppStart := strings.Index(pContent, "<w:pPr>"); ppStart >= 0 {
		if ppEnd := strings.Index(pContent[ppStart:], "</w:pPr>"); ppEnd >= 0 {
			pPr = pContent[ppStart : ppStart+ppEnd+len("</w:pPr>")]
		}
	}

	// Find the opening <w:tc> tag (may have attributes)
	tcOpenEnd := strings.Index(tcContent, ">")
	tcOpenTag := tcContent[:tcOpenEnd+1]

	// Find the opening <w:p> tag
	pOpenEnd := strings.Index(pContent, ">")
	pOpenTag := pContent[:pOpenEnd+1]

	var b strings.Builder
	b.WriteString(tcOpenTag)
	b.WriteString(tcPr)
	b.WriteString(pOpenTag)
	b.WriteString(pPr)
	if escaped != "" {
		b.WriteString(`<w:r><w:t xml:space="preserve">`)
		b.WriteString(escaped)
		b.WriteString(`</w:t></w:r>`)
	}
	b.WriteString("</w:p>")
	b.WriteString("</w:tc>")
	return b.String(), nil
}

// AppendTableRow clones the last row of the Nth table (1-based) and replaces
// each cell's text with the corresponding entry in cells. If cells has fewer
// entries than the row has columns, the remaining cells are cleared. If it has
// more entries, the extras are ignored.
func (u *Updater) AppendTableRow(tableIndex int, cells []string) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if tableIndex < 1 {
		return fmt.Errorf("tableIndex must be >= 1")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := appendTableRowContent(raw, tableIndex, cells)
	if err != nil {
		return err
	}

	return os.WriteFile(docPath, updated, 0o644)
}

// appendTableRowContent performs the XML surgery for AppendTableRow.
func appendTableRowContent(raw []byte, tableIndex int, cells []string) ([]byte, error) {
	content := string(raw)

	// Locate Nth <w:tbl> block.
	tblStart, tblEnd, err := findNthXMLBlock(content, "w:tbl", tableIndex)
	if err != nil {
		return nil, fmt.Errorf("table %d not found: %w", tableIndex, err)
	}
	tblContent := content[tblStart:tblEnd]

	// Count rows and grab the last one to use as a template.
	rowCount := 0
	lastTRStart, lastTREnd := 0, 0
	remaining := tblContent
	offset := 0
	for {
		s, e, rerr := findNthXMLBlock(remaining, "w:tr", rowCount+1)
		if rerr != nil {
			break
		}
		rowCount++
		lastTRStart = offset + s
		lastTREnd = offset + e
		advance := e
		offset += advance
		remaining = remaining[advance:]
	}
	if rowCount == 0 {
		return nil, fmt.Errorf("table %d has no rows", tableIndex)
	}

	templateRow := tblContent[lastTRStart:lastTREnd]

	// Count cells in the template row.
	cellCount := 0
	for {
		_, _, cerr := findNthXMLBlock(templateRow, "w:tc", cellCount+1)
		if cerr != nil {
			break
		}
		cellCount++
	}

	// Build the new row by cloning the template and replacing each cell's text.
	newRow := templateRow
	for i := 0; i < cellCount; i++ {
		var val string
		if i < len(cells) {
			val = cells[i]
		}
		tcS, tcE, cerr := findNthXMLBlock(newRow, "w:tc", i+1)
		if cerr != nil {
			break
		}
		updatedTC, rerr := replaceCellText(newRow[tcS:tcE], val)
		if rerr != nil {
			return nil, fmt.Errorf("replace cell %d text: %w", i+1, rerr)
		}
		newRow = newRow[:tcS] + updatedTC + newRow[tcE:]
	}

	// Insert the new row before </w:tbl>.
	closeTbl := "</w:tbl>"
	closePos := strings.LastIndex(tblContent, closeTbl)
	if closePos < 0 {
		return nil, fmt.Errorf("table %d: missing </w:tbl>", tableIndex)
	}
	newTbl := tblContent[:closePos] + newRow + tblContent[closePos:]
	result := content[:tblStart] + newTbl + content[tblEnd:]
	return []byte(result), nil
}

// InsertTableRowBefore inserts a new row immediately before the beforeRowIndex-th
// row (1-based) of the tableIndex-th table (1-based). The row at
// beforeRowIndex-1 (i.e. the row immediately before the gap) is used as a
// formatting template. cells values replace each cell's text content. If cells
// has fewer entries than the template row has columns, the remaining cells are
// cleared; extra entries are ignored.
func (u *Updater) InsertTableRowBefore(tableIndex, beforeRowIndex int, cells []string) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if tableIndex < 1 {
		return fmt.Errorf("tableIndex must be >= 1")
	}
	if beforeRowIndex < 1 {
		return fmt.Errorf("beforeRowIndex must be >= 1")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := insertTableRowBeforeContent(raw, tableIndex, beforeRowIndex, cells)
	if err != nil {
		return err
	}

	return os.WriteFile(docPath, updated, 0o644)
}

// insertTableRowBeforeContent performs the XML surgery for InsertTableRowBefore.
func insertTableRowBeforeContent(raw []byte, tableIndex, beforeRowIndex int, cells []string) ([]byte, error) {
	content := string(raw)

	// Locate Nth <w:tbl> block.
	tblStart, tblEnd, err := findNthXMLBlock(content, "w:tbl", tableIndex)
	if err != nil {
		return nil, fmt.Errorf("table %d not found: %w", tableIndex, err)
	}
	tblContent := content[tblStart:tblEnd]

	// Locate the row we will insert before — record its start offset in tblContent.
	insertTrStart, _, err := findNthXMLBlock(tblContent, "w:tr", beforeRowIndex)
	if err != nil {
		return nil, fmt.Errorf("table %d row %d not found: %w", tableIndex, beforeRowIndex, err)
	}

	// Use the row just before the insertion point as the formatting template.
	// If inserting before row 1, use row 1 itself as the template.
	templateIdx := beforeRowIndex - 1
	if templateIdx < 1 {
		templateIdx = 1
	}
	trS, trE, terr := findNthXMLBlock(tblContent, "w:tr", templateIdx)
	if terr != nil {
		return nil, fmt.Errorf("table %d template row %d not found: %w", tableIndex, templateIdx, terr)
	}
	templateRow := tblContent[trS:trE]

	// Count cells in the template row.
	cellCount := 0
	for {
		_, _, cerr := findNthXMLBlock(templateRow, "w:tc", cellCount+1)
		if cerr != nil {
			break
		}
		cellCount++
	}

	// Build the new row by cloning the template and replacing each cell's text.
	newRow := templateRow
	for i := 0; i < cellCount; i++ {
		var val string
		if i < len(cells) {
			val = cells[i]
		}
		tcS, tcE, cerr := findNthXMLBlock(newRow, "w:tc", i+1)
		if cerr != nil {
			break
		}
		updatedTC, rerr := replaceCellText(newRow[tcS:tcE], val)
		if rerr != nil {
			return nil, fmt.Errorf("replace cell %d text: %w", i+1, rerr)
		}
		newRow = newRow[:tcS] + updatedTC + newRow[tcE:]
	}

	// Insert the new row before insertTrStart in tblContent.
	newTbl := tblContent[:insertTrStart] + newRow + tblContent[insertTrStart:]
	result := content[:tblStart] + newTbl + content[tblEnd:]
	return []byte(result), nil
}

package godocx

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// MergeTableCellsHorizontal merges cells in a single row across columns.
// tableIndex, row, startCol, endCol are all 1-based.
// The content of the first cell (startCol) is preserved; merged cells are removed.
//
// Note: nested tables (a table inside a table cell) are not supported.
func (u *Updater) MergeTableCellsHorizontal(tableIndex, row, startCol, endCol int) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if tableIndex < 1 {
		return fmt.Errorf("tableIndex must be >= 1")
	}
	if row < 1 {
		return fmt.Errorf("row must be >= 1")
	}
	if startCol < 1 {
		return fmt.Errorf("startCol must be >= 1")
	}
	if endCol <= startCol {
		return fmt.Errorf("endCol must be greater than startCol")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := mergeTableCellsHorizontal(raw, tableIndex, row, startCol, endCol)
	if err != nil {
		return err
	}

	return atomicWriteFile(docPath, updated, 0o644)
}

// MergeTableCellsVertical merges cells in a single column across rows.
// tableIndex, startRow, endRow, col are all 1-based.
// The content of the first cell (startRow) is preserved; subsequent cells are marked
// as continuation cells (their content remains but Word displays only the first cell).
//
// Note: nested tables (a table inside a table cell) are not supported.
func (u *Updater) MergeTableCellsVertical(tableIndex, startRow, endRow, col int) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if tableIndex < 1 {
		return fmt.Errorf("tableIndex must be >= 1")
	}
	if col < 1 {
		return fmt.Errorf("col must be >= 1")
	}
	if startRow < 1 {
		return fmt.Errorf("startRow must be >= 1")
	}
	if endRow <= startRow {
		return fmt.Errorf("endRow must be greater than startRow")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := mergeTableCellsVertical(raw, tableIndex, startRow, endRow, col)
	if err != nil {
		return err
	}

	return atomicWriteFile(docPath, updated, 0o644)
}

// mergeTableCellsHorizontal performs horizontal cell merge on raw document XML.
func mergeTableCellsHorizontal(raw []byte, tableIndex, row, startCol, endCol int) ([]byte, error) {
	if endCol <= startCol {
		return nil, fmt.Errorf("endCol (%d) must be greater than startCol (%d)", endCol, startCol)
	}

	content := string(raw)
	span := endCol - startCol + 1

	tblStart, tblEnd, err := findNthXMLBlock(content, "w:tbl", tableIndex)
	if err != nil {
		return nil, fmt.Errorf("table %d not found: %w", tableIndex, err)
	}
	tblContent := content[tblStart:tblEnd]

	trStart, trEnd, err := findNthXMLBlock(tblContent, "w:tr", row)
	if err != nil {
		return nil, fmt.Errorf("row %d not found: %w", row, err)
	}
	trContent := tblContent[trStart:trEnd]

	tcStart1, tcEnd1, err := findNthXMLBlock(trContent, "w:tc", startCol)
	if err != nil {
		return nil, fmt.Errorf("cell %d not found: %w", startCol, err)
	}
	firstCell := trContent[tcStart1:tcEnd1]

	_, tcEndLast, err := findNthXMLBlock(trContent, "w:tc", endCol)
	if err != nil {
		return nil, fmt.Errorf("cell %d not found: %w", endCol, err)
	}

	modifiedCell := injectTcPrElement(firstCell, fmt.Sprintf(`<w:gridSpan w:val="%d"/>`, span))

	newTR := trContent[:tcStart1] + modifiedCell + trContent[tcEndLast:]
	newTbl := tblContent[:trStart] + newTR + tblContent[trEnd:]
	result := content[:tblStart] + newTbl + content[tblEnd:]

	return []byte(result), nil
}

// mergeTableCellsVertical performs vertical cell merge on raw document XML.
func mergeTableCellsVertical(raw []byte, tableIndex, startRow, endRow, col int) ([]byte, error) {
	if endRow <= startRow {
		return nil, fmt.Errorf("endRow (%d) must be greater than startRow (%d)", endRow, startRow)
	}

	result := string(raw)

	for row := startRow; row <= endRow; row++ {
		tblStart, tblEnd, err := findNthXMLBlock(result, "w:tbl", tableIndex)
		if err != nil {
			return nil, fmt.Errorf("table %d not found: %w", tableIndex, err)
		}
		tblContent := result[tblStart:tblEnd]

		trStart, trEnd, err := findNthXMLBlock(tblContent, "w:tr", row)
		if err != nil {
			return nil, fmt.Errorf("row %d not found: %w", row, err)
		}
		trContent := tblContent[trStart:trEnd]

		tcStart, tcEnd, err := findNthXMLBlock(trContent, "w:tc", col)
		if err != nil {
			return nil, fmt.Errorf("cell at row %d col %d not found: %w", row, col, err)
		}
		cellContent := trContent[tcStart:tcEnd]

		var mergeElement string
		if row == startRow {
			mergeElement = `<w:vMerge w:val="restart"/>`
		} else {
			mergeElement = `<w:vMerge/>`
		}

		modifiedCell := injectTcPrElement(cellContent, mergeElement)

		newTR := trContent[:tcStart] + modifiedCell + trContent[tcEnd:]
		newTbl := tblContent[:trStart] + newTR + tblContent[trEnd:]
		result = result[:tblStart] + newTbl + result[tblEnd:]
	}

	return []byte(result), nil
}

// injectTcPrElement injects an XML element into a table cell's <w:tcPr> block.
// If <w:tcPr> doesn't exist, one is created.
func injectTcPrElement(cellContent string, element string) string {
	if idx := strings.Index(cellContent, "<w:tcPr/>"); idx >= 0 {
		return cellContent[:idx] + "<w:tcPr>" + element + "</w:tcPr>" + cellContent[idx+len("<w:tcPr/>"):]
	}

	if idx := strings.Index(cellContent, "<w:tcPr>"); idx >= 0 {
		insertPos := idx + len("<w:tcPr>")
		return cellContent[:insertPos] + element + cellContent[insertPos:]
	}
	if idx := strings.Index(cellContent, "<w:tcPr "); idx >= 0 {
		closeIdx := strings.Index(cellContent[idx:], ">")
		if closeIdx >= 0 {
			insertPos := idx + closeIdx + 1
			return cellContent[:insertPos] + element + cellContent[insertPos:]
		}
	}

	openEnd := strings.Index(cellContent, ">")
	if openEnd < 0 {
		return cellContent
	}
	insertPos := openEnd + 1
	return cellContent[:insertPos] + "<w:tcPr>" + element + "</w:tcPr>" + cellContent[insertPos:]
}

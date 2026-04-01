package godocx_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

// extractGridColumnWidths extracts grid column widths from document XML
func extractGridColumnWidths(docXML string) []int {
	// Find the LAST table's grid, not the first
	// This is important for tests that insert tables into templates with existing tables
	lastTblGridStart := strings.LastIndex(docXML, "<w:tblGrid>")
	if lastTblGridStart == -1 {
		return nil
	}

	tblGridEnd := strings.Index(docXML[lastTblGridStart:], "</w:tblGrid>")
	if tblGridEnd == -1 {
		return nil
	}
	tblGridEnd += lastTblGridStart

	gridXML := docXML[lastTblGridStart : tblGridEnd+len("</w:tblGrid>")]
	return extractGridColumns(gridXML)
}

// extractGridColumns helper to parse XML and find grid columns
func extractGridColumns(tblGridXML string) []int {
	var widths []int
	start := 0
	for {
		idx := strings.Index(tblGridXML[start:], `<w:gridCol w:w="`)
		if idx == -1 {
			break
		}
		start = start + idx + len(`<w:gridCol w:w="`)
		endIdx := strings.Index(tblGridXML[start:], `"`)
		if endIdx == -1 {
			break
		}
		widthStr := tblGridXML[start : start+endIdx]
		var width int
		_, err := parseWidth(widthStr, &width)
		if err != nil {
			break
		}
		widths = append(widths, width)
		start = start + endIdx
	}
	return widths
}

// parseWidth parses width string to int
func parseWidth(s string, dst *int) (int, error) {
	*dst = 0
	for _, c := range s {
		if c >= '0' && c <= '9' {
			*dst = *dst*10 + int(c-'0')
		}
	}
	return *dst, nil
}

// TestTableProportionalColumnWidths verifies proportional sizing based on content
func TestTableProportionalColumnWidths(t *testing.T) {
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

	// Create table with proportional widths
	// Column 1: "ID" (2 chars) - shortest
	// Column 2: "Description" (11 chars) - longest
	// Column 3: "Price" (5 chars) - medium
	// Proportions should be approximately 2:11:5
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "ID"},
			{Title: "Description"},
			{Title: "Price"},
		},
		Rows: [][]string{
			{"1", "Short desc", "$9.99"},
			{"2", "A much longer description here", "$19.99"},
			{"3", "Medium length text", "$14.99"},
		},
		HeaderBold:               true,
		ProportionalColumnWidths: true,
		// Column widths should be proportional to content length
		// ID (2) : Description (31) : Price (7) = roughly 1:15:3
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	gridColumns := extractGridColumnWidths(docXML)

	if len(gridColumns) != 3 {
		t.Fatalf("Expected 3 grid columns, got %d", len(gridColumns))
	}

	// Total width should be 9360 (default 100% of Letter portrait)
	totalWidth := gridColumns[0] + gridColumns[1] + gridColumns[2]
	expectedTotal := 9360
	if totalWidth != expectedTotal {
		t.Errorf("Total grid width = %d, want %d", totalWidth, expectedTotal)
	}

	// Description column should be wider than ID and Price
	if gridColumns[1] <= gridColumns[0] || gridColumns[1] <= gridColumns[2] {
		t.Errorf("Description column (width: %d) should be widest. Got ID: %d, Description: %d, Price: %d",
			gridColumns[1], gridColumns[0], gridColumns[1], gridColumns[2])
	}

	// ID column should be narrowest
	if gridColumns[0] >= gridColumns[2] {
		t.Errorf("ID column (width: %d) should be narrowest. Got ID: %d, Price: %d",
			gridColumns[0], gridColumns[0], gridColumns[2])
	}

	t.Logf("Proportional widths - ID: %d, Description: %d, Price: %d", gridColumns[0], gridColumns[1], gridColumns[2])
}

// TestTableProportionalWithFixedWidth verifies proportional sizing with fixed table width
func TestTableProportionalWithFixedWidth(t *testing.T) {
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

	// Fixed width table (6 inches = 8640 twips) with proportional columns
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "A"},
			{Title: "Much Longer Header"},
			{Title: "B"},
		},
		Rows: [][]string{
			{"X", "Y content here", "Z"},
		},
		HeaderBold:               true,
		ProportionalColumnWidths: true,
		TableWidthType:           godocx.TableWidthFixed,
		TableWidth:               8640, // 6 inches
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	gridColumns := extractGridColumnWidths(docXML)

	if len(gridColumns) != 3 {
		t.Fatalf("Expected 3 grid columns, got %d", len(gridColumns))
	}

	// Total should equal fixed width
	totalWidth := gridColumns[0] + gridColumns[1] + gridColumns[2]
	if totalWidth != 8640 {
		t.Errorf("Total grid width = %d, want 8640", totalWidth)
	}

	// Middle column should be wider (longer header)
	if gridColumns[1] <= gridColumns[0] || gridColumns[1] <= gridColumns[2] {
		t.Errorf("Middle column should be widest; got A: %d, Middle: %d, B: %d",
			gridColumns[0], gridColumns[1], gridColumns[2])
	}

	t.Logf("Fixed proportional widths - A: %d, Middle: %d, B: %d", gridColumns[0], gridColumns[1], gridColumns[2])
}

// TestTableProportionalWithPercentageWidth verifies proportional with percentage mode
func TestTableProportionalWithPercentageWidth(t *testing.T) {
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

	// 50% width with proportional columns
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Short"},
			{Title: "Very Long Column Header"},
			{Title: "Med"},
		},
		Rows: [][]string{
			{"S", "VeryLongContentHere", "M"},
		},
		HeaderBold:               true,
		ProportionalColumnWidths: true,
		TableWidthType:           godocx.TableWidthPercentage,
		TableWidth:               2500, // 50%
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	gridColumns := extractGridColumnWidths(docXML)

	if len(gridColumns) != 3 {
		t.Fatalf("Expected 3 grid columns, got %d", len(gridColumns))
	}

	// Total should be 50% of available width: (9360 * 2500) / 5000 = 4680
	totalWidth := gridColumns[0] + gridColumns[1] + gridColumns[2]
	expectedTotal := 4680
	if totalWidth != expectedTotal {
		t.Errorf("Total grid width = %d, want %d (50%% of 9360)", totalWidth, expectedTotal)
	}

	t.Logf("Percentage proportional widths (50%%) - Short: %d, Long: %d, Med: %d", gridColumns[0], gridColumns[1], gridColumns[2])
}

// TestTableEqualWidthsStillDefault verifies equal widths when ProportionalColumnWidths is false
func TestTableEqualWidthsStillDefault(t *testing.T) {
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

	// Default: equal widths
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "X"},
			{Title: "Very Long Header Text"},
			{Title: "Y"},
		},
		Rows: [][]string{
			{"1", "2", "3"},
		},
		HeaderBold: true,
		// ProportionalColumnWidths: false (default)
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	gridColumns := extractGridColumnWidths(docXML)

	if len(gridColumns) != 3 {
		t.Fatalf("Expected 3 grid columns, got %d", len(gridColumns))
	}

	// All columns should be equal width: 9360 / 3 = 3120
	expectedWidth := 3120
	for i, width := range gridColumns {
		if width != expectedWidth {
			t.Errorf("Column %d width = %d, want %d (equal distribution)", i+1, width, expectedWidth)
		}
	}

	t.Logf("Equal widths (default) - All columns: %d twips", expectedWidth)
}

// TestTableProportionalIgnoredWithExplicitWidths verifies explicit widths take precedence
func TestTableProportionalIgnoredWithExplicitWidths(t *testing.T) {
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

	// Explicit column widths should ignore ProportionalColumnWidths
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "A"},
			{Title: "Very Long Header"},
			{Title: "B"},
		},
		ColumnWidths: []int{1000, 2000, 1000}, // Explicit widths
		Rows: [][]string{
			{"1", "2", "3"},
		},
		HeaderBold:               true,
		ProportionalColumnWidths: true, // Should be ignored
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	gridColumns := extractGridColumnWidths(docXML)

	// Should use explicit widths, not proportional
	expectedWidths := []int{1000, 2000, 1000}
	for i, expected := range expectedWidths {
		if gridColumns[i] != expected {
			t.Errorf("Column %d width = %d, want %d (explicit takes precedence)", i+1, gridColumns[i], expected)
		}
	}

	t.Logf("Explicit widths take precedence: %v", gridColumns)
}

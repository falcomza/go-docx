package docxupdater_test

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

func TestInsertBasicTable(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create a simple table
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Name"},
			{Title: "Age"},
			{Title: "City"},
		},
		Rows: [][]string{
			{"Alice", "30", "New York"},
			{"Bob", "25", "Los Angeles"},
			{"Charlie", "35", "Chicago"},
		},
		HeaderBold: true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify table was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:tbl>") {
		t.Error("Table element not found in document.xml")
	}
	if !strings.Contains(docXML, "Name") || !strings.Contains(docXML, "Alice") {
		t.Error("Table content not found in document.xml")
	}
}

func TestInsertTableWithStyling(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Product", Alignment: docxupdater.CellAlignLeft},
			{Title: "Price", Alignment: docxupdater.CellAlignRight},
			{Title: "Stock", Alignment: docxupdater.CellAlignCenter},
		},
		Rows: [][]string{
			{"Laptop", "$999", "15"},
			{"Mouse", "$29", "50"},
			{"Keyboard", "$79", "30"},
		},
		HeaderBold:        true,
		HeaderBackground:  "4472C4",
		HeaderAlignment:   docxupdater.CellAlignCenter,
		AlternateRowColor: "F2F2F2",
		BorderStyle:       docxupdater.BorderSingle,
		BorderSize:        6,
		TableAlignment:    docxupdater.AlignCenter,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "4472C4") {
		t.Error("Header background color not found")
	}
	if !strings.Contains(docXML, "F2F2F2") {
		t.Error("Alternate row color not found")
	}
}

func TestInsertTableWithRepeatHeader(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with many rows to test repeat header
	rows := make([][]string, 50)
	for i := range rows {
		rows[i] = []string{
			fmt.Sprintf("Item %d", i+1),
			fmt.Sprintf("Description for item %d", i+1),
			fmt.Sprintf("$%d.00", (i+1)*10),
		}
	}

	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Item"},
			{Title: "Description"},
			{Title: "Price"},
		},
		Rows:             rows,
		HeaderBold:       true,
		RepeatHeader:     true,
		HeaderBackground: "2E75B5",
		TableAlignment:   docxupdater.AlignCenter,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:tblHeader/>") {
		t.Error("Table header repeat property not found")
	}
}

func TestInsertTableInvalidRows(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Try to create table with mismatched column count
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Col1"},
			{Title: "Col2"},
		},
		Rows: [][]string{
			{"A", "B"},
			{"C", "D", "E"}, // Wrong number of cells
		},
	})
	if err == nil {
		t.Error("Expected error for mismatched column count, got nil")
	}
}

func TestInsertTableNoColumns(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns:  []docxupdater.ColumnDefinition{},
		Rows:     [][]string{},
	})
	if err == nil {
		t.Error("Expected error for no columns, got nil")
	}
}

func TestInsertTableCustomWidths(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Code"},
			{Title: "Description"},
			{Title: "Status"},
		},
		ColumnWidths: []int{1000, 3000, 1000}, // Custom widths in twips
		Rows: [][]string{
			{"A1", "First item with long description", "Active"},
			{"B2", "Second item", "Pending"},
		},
		HeaderBold: true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:w="1000"`) || !strings.Contains(docXML, `w:w="3000"`) {
		t.Error("Custom column widths not found in document")
	}
}

func TestInsertTableWidthPercentage(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with 100% width (default)
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		HeaderBold: true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	// Should default to 100% width (5000 in pct mode)
	if !strings.Contains(docXML, `w:type="pct"`) || !strings.Contains(docXML, `w:w="5000"`) {
		t.Error("Default 100% width not found in document")
	}
}

func TestInsertTableWidthFixed(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with fixed width in twips
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		TableWidthType: docxupdater.TableWidthFixed,
		TableWidth:     7200, // 5 inches
		HeaderBold:     true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:type="dxa"`) || !strings.Contains(docXML, `w:w="7200"`) {
		t.Error("Fixed width not found in document")
	}
}

func TestInsertTableWidthAuto(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with auto width
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		TableWidthType: docxupdater.TableWidthAuto,
		HeaderBold:     true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:type="auto"`) {
		t.Error("Auto width not found in document")
	}
}

func TestInsertTableWidth50Percent(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with 50% width
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		TableWidthType: docxupdater.TableWidthPercentage,
		TableWidth:     2500, // 50% (5000 = 100%)
		HeaderBold:     true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:type="pct"`) || !strings.Contains(docXML, `w:w="2500"`) {
		t.Error("50% width not found in document")
	}
}

func TestInsertTableRowHeightExact(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with exact row heights
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Name"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"Row 1", "Data 1"},
			{"Row 2", "Data 2"},
		},
		HeaderRowHeight:  720, // 0.5 inch for header
		HeaderHeightRule: docxupdater.RowHeightExact,
		RowHeight:        360, // 0.25 inch for data rows
		RowHeightRule:    docxupdater.RowHeightExact,
		HeaderBold:       true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:val="720"`) || !strings.Contains(docXML, `w:hRule="exact"`) {
		t.Error("Exact row height not found in document")
	}
	if !strings.Contains(docXML, `w:val="360"`) {
		t.Error("Data row height not found in document")
	}
}

func TestInsertTableRowHeightAtLeast(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with minimum row heights
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Description"},
		},
		Rows: [][]string{
			{"This is a row that might have varying content length"},
		},
		RowHeight:     500,
		RowHeightRule: docxupdater.RowHeightAtLeast,
		HeaderBold:    true,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:val="500"`) || !strings.Contains(docXML, `w:hRule="atLeast"`) {
		t.Error("AtLeast row height not found in document")
	}
}

func TestInsertTableRowHeightAuto(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with auto row heights (default)
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Auto Height"},
		},
		Rows: [][]string{
			{"Content determines height"},
		},
		HeaderBold: true,
		// RowHeightRule defaults to RowHeightAuto
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// With auto height (default), no w:trHeight should be present
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	// Auto height means no height specification in XML
	if strings.Count(docXML, `<w:trHeight`) > 0 {
		t.Error("Auto height should not have w:trHeight element")
	}
}

func TestInsertTableWithNamedStyles(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with named Word styles
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Heading Column"},
			{Title: "Normal Column"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
			{"Data 3", "Data 4"},
		},
		HeaderStyleName:  "Heading1", // Word style for header
		RowStyleName:     "BodyText", // Word style for data rows
		HeaderBold:       true,
		HeaderBackground: "4472C4",
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:pStyle w:val="Heading1"/>`) {
		t.Error("Header paragraph style not found in document")
	}
	if !strings.Contains(docXML, `<w:pStyle w:val="BodyText"/>`) {
		t.Error("Row paragraph style not found in document")
	}
}

func TestInsertTableMixedDirectAndNamedStyles(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table mixing named styles and direct formatting
	err = u.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		HeaderStyleName:  "Heading2", // Named style
		HeaderBold:       true,       // Plus direct formatting
		HeaderBackground: "70AD47",
		RowStyleName:     "Normal", // Named style for rows
		RowStyle: docxupdater.CellStyle{ // Plus direct formatting
			FontSize: 20, // 10pt
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	// Should have both style reference and direct formatting
	if !strings.Contains(docXML, `<w:pStyle w:val="Heading2"/>`) {
		t.Error("Header paragraph style not found")
	}
	if !strings.Contains(docXML, `<w:pStyle w:val="Normal"/>`) {
		t.Error("Row paragraph style not found")
	}
	if !strings.Contains(docXML, `<w:b/>`) {
		t.Error("Direct bold formatting not found")
	}
}

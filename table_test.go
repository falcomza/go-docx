package godocx_test

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

func TestInsertBasicTable(t *testing.T) {
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

	// Create a simple table
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Product", Alignment: godocx.CellAlignLeft},
			{Title: "Price", Alignment: godocx.CellAlignRight},
			{Title: "Stock", Alignment: godocx.CellAlignCenter},
		},
		Rows: [][]string{
			{"Laptop", "$999", "15"},
			{"Mouse", "$29", "50"},
			{"Keyboard", "$79", "30"},
		},
		HeaderBold:        true,
		HeaderBackground:  "4472C4",
		HeaderAlignment:   godocx.CellAlignCenter,
		AlternateRowColor: "F2F2F2",
		BorderStyle:       godocx.BorderSingle,
		BorderSize:        6,
		TableAlignment:    godocx.AlignCenter,
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

	u, err := godocx.New(inputPath)
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

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Item"},
			{Title: "Description"},
			{Title: "Price"},
		},
		Rows:             rows,
		HeaderBold:       true,
		RepeatHeader:     true,
		HeaderBackground: "2E75B5",
		TableAlignment:   godocx.AlignCenter,
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Try to create table with mismatched column count
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns:  []godocx.ColumnDefinition{},
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with 100% width (default)
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with fixed width in twips
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		TableWidthType: godocx.TableWidthFixed,
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with auto width
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		TableWidthType: godocx.TableWidthAuto,
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with 50% width
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Column 1"},
			{Title: "Column 2"},
		},
		Rows: [][]string{
			{"Data 1", "Data 2"},
		},
		TableWidthType: godocx.TableWidthPercentage,
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with exact row heights
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Name"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"Row 1", "Data 1"},
			{"Row 2", "Data 2"},
		},
		HeaderRowHeight:  720, // 0.5 inch for header
		HeaderHeightRule: godocx.RowHeightExact,
		RowHeight:        360, // 0.25 inch for data rows
		RowHeightRule:    godocx.RowHeightExact,
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with minimum row heights
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Description"},
		},
		Rows: [][]string{
			{"This is a row that might have varying content length"},
		},
		RowHeight:     500,
		RowHeightRule: godocx.RowHeightAtLeast,
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with auto row heights (default)
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table with named Word styles
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create table mixing named styles and direct formatting
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
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
		RowStyle: godocx.CellStyle{ // Plus direct formatting
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

func TestInsertTableWithConditionalCellColors(t *testing.T) {
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

	// Create table with conditional cell coloring based on status
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Service"},
			{Title: "Status"},
			{Title: "Uptime"},
		},
		Rows: [][]string{
			{"Database", "Critical", "45%"},
			{"API", "Normal", "99.9%"},
			{"Cache", "Warning", "85%"},
			{"Auth", "Normal", "99.5%"},
		},
		HeaderBold:       true,
		HeaderBackground: "4472C4",
		RowStyle: godocx.CellStyle{
			FontSize: 20,
		},
		// Define conditional styles for status values
		ConditionalStyles: map[string]godocx.CellStyle{
			"Critical": {
				Background: "FF0000", // Red
				FontColor:  "FFFFFF", // White text
				Bold:       true,
			},
			"Warning": {
				Background: "FFA500", // Orange
				FontColor:  "000000",
			},
			"Normal": {
				Background: "00B050", // Green
				FontColor:  "FFFFFF",
			},
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Verify Critical status has red background
	if !strings.Contains(docXML, `w:fill="FF0000"`) {
		t.Error("Critical red background not found")
	}

	// Verify Warning status has orange background
	if !strings.Contains(docXML, `w:fill="FFA500"`) {
		t.Error("Warning orange background not found")
	}

	// Verify Normal status has green background
	if !strings.Contains(docXML, `w:fill="00B050"`) {
		t.Error("Normal green background not found")
	}

	// Verify Critical has white font color
	if !strings.Contains(docXML, `<w:color w:val="FFFFFF"/>`) {
		t.Error("White font color for Critical not found")
	}

	// Verify header background is still applied
	if !strings.Contains(docXML, `w:fill="4472C4"`) {
		t.Error("Header background color not found")
	}
}

func TestInsertTableConditionalCaseInsensitive(t *testing.T) {
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

	// Test case-insensitive matching
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Item"},
			{Title: "Priority"},
		},
		Rows: [][]string{
			{"Task 1", "HIGH"},     // Uppercase
			{"Task 2", "high"},     // Lowercase
			{"Task 3", "High"},     // Mixed case
			{"Task 4", "  High  "}, // With spaces
		},
		ConditionalStyles: map[string]godocx.CellStyle{
			"High": {
				Background: "FF6B6B",
			},
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Count occurrences of the conditional background color
	// All 4 rows should have the same background
	count := strings.Count(docXML, `w:fill="FF6B6B"`)
	if count < 4 {
		t.Errorf("Expected at least 4 occurrences of conditional color, got %d", count)
	}
}

func TestInsertTableConditionalWithRowStyle(t *testing.T) {
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

	// Test that conditional styles override row styles
	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Metric"},
			{Title: "Rating"},
		},
		Rows: [][]string{
			{"CPU", "Good"},
			{"Memory", "Poor"},
			{"Disk", "Good"},
		},
		RowStyle: godocx.CellStyle{
			Background: "E7E6E6", // Gray default background
			FontSize:   20,
		},
		// Conditional styles should override the gray background
		ConditionalStyles: map[string]godocx.CellStyle{
			"Good": {
				Background: "00B050", // Green
			},
			"Poor": {
				Background: "FF0000", // Red
				FontColor:  "FFFFFF",
			},
		},
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Verify conditional colors are present
	if !strings.Contains(docXML, `w:fill="00B050"`) {
		t.Error("Good green background not found")
	}
	if !strings.Contains(docXML, `w:fill="FF0000"`) {
		t.Error("Poor red background not found")
	}

	// Verify default gray background is also present (for non-matching cells)
	if !strings.Contains(docXML, `w:fill="E7E6E6"`) {
		t.Error("Default gray background not found for non-matching cells")
	}
}

// TestInsertTableDefaultStylesInjectedIntoStylesXML verifies that the default
// paragraph styles "Table Header" (header cells) and "Table" (data cells) are
// automatically injected into styles.xml so LibreOffice/Word can resolve them.
// Without this, both applications collapse the table to unstyled plain text.
func TestInsertTableDefaultStylesInjectedIntoStylesXML(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Name"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"alpha", "1"},
			{"beta", "2"},
		},
	})
	if err != nil {
		t.Fatalf("InsertTable: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// styles.xml must exist and must define both default style IDs.
	stylesXML := readZipEntry(t, outputPath, "word/styles.xml")
	if stylesXML == "" {
		t.Fatal("word/styles.xml is absent from output DOCX")
	}
	for _, id := range []string{"Table Header", "Table"} {
		attr := fmt.Sprintf(`w:styleId="%s"`, id)
		if !strings.Contains(stylesXML, attr) {
			t.Errorf("styles.xml missing style definition for %q (looked for %s)", id, attr)
		}
	}

	// document.xml must still reference those styles in cells.
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:pStyle w:val="Table Header"/>`) {
		t.Error("document.xml missing 'Table Header' pStyle on header cells")
	}
	if !strings.Contains(docXML, `<w:pStyle w:val="Table"/>`) {
		t.Error("document.xml missing 'Table' pStyle on data cells")
	}
}

// TestInsertTableCustomStyleNamesAutoInjected verifies that caller-supplied
// custom style names that don't exist in styles.xml are also auto-injected.
func TestInsertTableCustomStyleNamesAutoInjected(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Col"},
		},
		Rows: [][]string{
			{"row1"},
		},
		HeaderStyleName: "MyCustomHeader",
		RowStyleName:    "MyCustomRow",
	})
	if err != nil {
		t.Fatalf("InsertTable: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	stylesXML := readZipEntry(t, outputPath, "word/styles.xml")
	for _, id := range []string{"MyCustomHeader", "MyCustomRow"} {
		attr := fmt.Sprintf(`w:styleId="%s"`, id)
		if !strings.Contains(stylesXML, attr) {
			t.Errorf("styles.xml missing auto-injected style %q", id)
		}
	}
}

// TestInsertTableNoInjectionWhenStylesExist verifies that if a style ID is
// already present in styles.xml it is not duplicated.
func TestInsertTableNoInjectionWhenStylesExist(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	// Build a fixture that already has "Table Header" and "Table" defined.
	docx := buildFixtureDocxWithStyles(t, []string{"Table Header", "Table"})
	if err := os.WriteFile(inputPath, docx, 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	if err := u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns:  []godocx.ColumnDefinition{{Title: "X"}},
		Rows:     [][]string{{"y"}},
	}); err != nil {
		t.Fatalf("InsertTable: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	stylesXML := readZipEntry(t, outputPath, "word/styles.xml")
	// Each style ID must appear exactly once (no duplicates).
	for _, id := range []string{"Table Header", "Table"} {
		attr := fmt.Sprintf(`w:styleId="%s"`, id)
		count := strings.Count(stylesXML, attr)
		if count != 1 {
			t.Errorf("style %q appears %d times in styles.xml, want exactly 1", id, count)
		}
	}
}

// buildFixtureDocxWithStyles builds a minimal DOCX that already contains
// paragraph style definitions for each name in styleIDs.
func buildFixtureDocxWithStyles(t *testing.T, styleIDs []string) []byte {
	t.Helper()

	var styleEntries strings.Builder
	for _, id := range styleIDs {
		styleEntries.WriteString(fmt.Sprintf(
			`<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="%s"/></w:style>`,
			id, id,
		))
	}
	stylesXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		styleEntries.String() +
		`</w:styles>`

	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZipEntry(t, zw, "[Content_Types].xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`+
			`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`+
			`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`+
			`</Types>`)
	addZipEntry(t, zw, "word/document.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`+
			`<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>`+
			`</w:document>`)
	addZipEntry(t, zw, "word/styles.xml", stylesXML)
	addZipEntry(t, zw, "word/_rels/document.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`+
			`</Relationships>`)
	if err := zw.Close(); err != nil {
		t.Fatalf("close zip: %v", err)
	}
	return buf.Bytes()
}

func buildFixtureDocxForTableUpdate(t *testing.T) *godocx.Updater {
	t.Helper()
	// Write a minimal DOCX to a temp file
	path := filepath.Join(t.TempDir(), "fixture.docx")
	if err := writeMinimalDocx(path); err != nil {
		t.Fatalf("writeMinimalDocx: %v", err)
	}
	u, err := godocx.New(path)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	// Insert a table with header row + 2 data rows
	err = u.InsertTable(godocx.TableOptions{
		Columns:  []godocx.ColumnDefinition{{Title: "Name"}, {Title: "Value"}},
		Rows:     [][]string{{"Alpha", "100"}, {"Beta", "200"}},
		Position: godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertTable: %v", err)
	}
	return u
}

func writeMinimalDocx(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	w := zip.NewWriter(f)
	entries := map[string]string{
		"[Content_Types].xml":          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
		"word/document.xml":            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`,
		"word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
		"_rels/.rels":                  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
	}
	for name, content := range entries {
		e, err := w.Create(name)
		if err != nil {
			return err
		}
		if _, err := e.Write([]byte(content)); err != nil {
			return err
		}
	}
	return w.Close()
}

func TestUpdateTableCell(t *testing.T) {
	u := buildFixtureDocxForTableUpdate(t)
	defer u.Cleanup()

	// Update row 2, col 2 (1-based, row 1 is header)
	if err := u.UpdateTableCell(1, 2, 2, "999"); err != nil {
		t.Fatalf("UpdateTableCell: %v", err)
	}

	tables, err := u.GetTableText()
	if err != nil {
		t.Fatalf("GetTableText: %v", err)
	}
	// tables[0] is first table; row index 1 is second row; col index 1 is second col
	got := tables[0][1][1]
	if got != "999" {
		t.Errorf("expected '999', got %q", got)
	}
}

func TestUpdateTableCellSpecialChars(t *testing.T) {
	u := buildFixtureDocxForTableUpdate(t)
	defer u.Cleanup()

	if err := u.UpdateTableCell(1, 1, 1, "A & B < C > D"); err != nil {
		t.Fatalf("UpdateTableCell: %v", err)
	}

	tables, err := u.GetTableText()
	if err != nil {
		t.Fatalf("GetTableText: %v", err)
	}
	got := tables[0][0][0]
	if got != "A & B < C > D" {
		t.Errorf("expected 'A & B < C > D', got %q", got)
	}
}

func TestUpdateTableCellValidation(t *testing.T) {
	u := buildFixtureDocxForTableUpdate(t)
	defer u.Cleanup()

	tests := []struct {
		tableIndex, row, col int
		wantErr              bool
	}{
		{0, 1, 1, true},  // tableIndex < 1
		{1, 0, 1, true},  // row < 1
		{1, 1, 0, true},  // col < 1
		{99, 1, 1, true}, // table doesn't exist
		{1, 99, 1, true}, // row doesn't exist
		{1, 1, 99, true}, // col doesn't exist
		{1, 1, 1, false}, // valid
	}
	for _, tt := range tests {
		err := u.UpdateTableCell(tt.tableIndex, tt.row, tt.col, "x")
		if (err != nil) != tt.wantErr {
			t.Errorf("UpdateTableCell(%d,%d,%d): got err=%v, wantErr=%v",
				tt.tableIndex, tt.row, tt.col, err, tt.wantErr)
		}
	}
}

func TestAppendTableRow(t *testing.T) {
	u := buildFixtureDocxForTableUpdate(t)
	defer u.Cleanup()

	// Table has header + 2 data rows = 3 rows before append.
	if err := u.AppendTableRow(1, []string{"Gamma", "300"}); err != nil {
		t.Fatalf("AppendTableRow: %v", err)
	}

	tables, err := u.GetTableText()
	if err != nil {
		t.Fatalf("GetTableText: %v", err)
	}
	tbl := tables[0]
	if len(tbl) != 4 {
		t.Fatalf("expected 4 rows after append, got %d", len(tbl))
	}
	last := tbl[len(tbl)-1]
	if last[0] != "Gamma" || last[1] != "300" {
		t.Errorf("last row = %v, want [Gamma 300]", last)
	}
}

func TestAppendTableRowFewerCells(t *testing.T) {
	u := buildFixtureDocxForTableUpdate(t)
	defer u.Cleanup()

	// Provide only one cell value for a 2-column table.
	if err := u.AppendTableRow(1, []string{"Delta"}); err != nil {
		t.Fatalf("AppendTableRow: %v", err)
	}

	tables, err := u.GetTableText()
	if err != nil {
		t.Fatalf("GetTableText: %v", err)
	}
	tbl := tables[0]
	last := tbl[len(tbl)-1]
	if last[0] != "Delta" {
		t.Errorf("cell[0] = %q, want Delta", last[0])
	}
	if last[1] != "" {
		t.Errorf("cell[1] = %q, want empty", last[1])
	}
}

func TestAppendTableRowValidation(t *testing.T) {
	u := buildFixtureDocxForTableUpdate(t)
	defer u.Cleanup()

	if err := u.AppendTableRow(0, []string{"x"}); err == nil {
		t.Error("expected error for tableIndex=0")
	}
	if err := u.AppendTableRow(99, []string{"x"}); err == nil {
		t.Error("expected error for non-existent table")
	}
}

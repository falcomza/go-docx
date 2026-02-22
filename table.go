package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// TableStyle defines table styling options
type TableStyle string

const (
	TableStyleGrid         TableStyle = "TableGrid"
	TableStyleGridAccent1  TableStyle = "LightShading-Accent1"
	TableStyleGridAccent2  TableStyle = "MediumShading1-Accent1"
	TableStylePlain        TableStyle = "TableNormal"
	TableStyleColorful     TableStyle = "ColorfulGrid-Accent1"
	TableStyleProfessional TableStyle = "LightGrid-Accent1"
)

// TableAlignment defines table alignment
type TableAlignment string

const (
	AlignLeft   TableAlignment = "left"
	AlignCenter TableAlignment = "center"
	AlignRight  TableAlignment = "right"
)

// CellAlignment defines cell content alignment
type CellAlignment string

const (
	CellAlignLeft   CellAlignment = "start"
	CellAlignCenter CellAlignment = "center"
	CellAlignRight  CellAlignment = "end"
)

// VerticalAlignment defines vertical alignment in cells
type VerticalAlignment string

const (
	VerticalAlignTop    VerticalAlignment = "top"
	VerticalAlignCenter VerticalAlignment = "center"
	VerticalAlignBottom VerticalAlignment = "bottom"
)

// BorderStyle defines table border style
type BorderStyle string

const (
	BorderSingle BorderStyle = "single"
	BorderDouble BorderStyle = "double"
	BorderDashed BorderStyle = "dashed"
	BorderDotted BorderStyle = "dotted"
	BorderNone   BorderStyle = "none"
)

// TableWidthType defines how table width is specified
type TableWidthType string

const (
	TableWidthAuto       TableWidthType = "auto" // Auto-fit to content
	TableWidthPercentage TableWidthType = "pct"  // Percentage of available width (5000 = 100%)
	TableWidthFixed      TableWidthType = "dxa"  // Fixed width in twips
)

// RowHeightRule defines how row height is interpreted
type RowHeightRule string

const (
	RowHeightAuto    RowHeightRule = "auto"    // Auto height based on content (default)
	RowHeightAtLeast RowHeightRule = "atLeast" // Minimum height, can grow
	RowHeightExact   RowHeightRule = "exact"   // Fixed height, no growth
)

// TableOptions defines comprehensive options for table creation
type TableOptions struct {
	// Position where to insert the table
	Position InsertPosition
	Anchor   string // Text anchor for relative positioning

	// Column definitions
	Columns      []ColumnDefinition // Column titles and properties
	ColumnWidths []int              // Optional: column widths in twips (1/1440 inch), nil for auto

	// Data rows (excluding header)
	Rows [][]string // Each inner slice is a row of cell data

	// ProportionalColumnWidths enables content-based proportional sizing
	// When true, column widths are calculated based on the length of content
	// (headers + longest cell) in each column. Wider content gets wider columns.
	// If false (default), all columns get equal width.
	// Ignored if ColumnWidths is explicitly specified.
	ProportionalColumnWidths bool

	// AvailableWidth specifies the usable page width in twips (page width - left margin - right margin)
	// Used for auto-calculating column widths in percentage mode
	// If 0 (default), uses standard Letter page width calculation: 12240 - 1440 - 1440 = 9360
	// Automatically computed if not set
	AvailableWidth int

	// Header styling
	HeaderStyle      CellStyle     // Style for header row
	HeaderStyleName  string        // Named Word style for header paragraphs (e.g., "Heading 1")
	RepeatHeader     bool          // Repeat header row on each page
	HeaderBackground string        // Hex color for header background (e.g., "4472C4")
	HeaderBold       bool          // Make header text bold
	HeaderAlignment  CellAlignment // Header text alignment

	// Row styling
	RowStyle          CellStyle         // Style for data rows
	RowStyleName      string            // Named Word style for data row paragraphs (e.g., "Normal")
	AlternateRowColor string            // Hex color for alternate rows (e.g., "F2F2F2")
	RowAlignment      CellAlignment     // Default cell alignment for data rows
	VerticalAlign     VerticalAlignment // Vertical alignment in cells

	// Row height
	HeaderRowHeight  int           // Header row height in twips, 0 for auto
	HeaderHeightRule RowHeightRule // Header height rule (auto, atLeast, exact)
	RowHeight        int           // Data row height in twips, 0 for auto
	RowHeightRule    RowHeightRule // Data row height rule (auto, atLeast, exact)

	// Table properties
	TableAlignment TableAlignment // Table alignment on page
	TableWidthType TableWidthType // Width type: auto, percentage, or fixed (default: percentage)
	TableWidth     int            // Width value: 0 for auto, 5000 for 100% (pct mode), or twips (dxa mode)
	TableStyle     TableStyle     // Predefined table style

	// Border properties
	BorderStyle BorderStyle // Border style

	// Caption options (nil for no caption)
	Caption     *CaptionOptions
	BorderSize  int    // Border width in eighths of a point (default: 4 = 0.5pt)
	BorderColor string // Hex color for borders (default: "000000")

	// Cell properties
	CellPadding int  // Cell padding in twips (default: 108 = 0.075")
	AutoFit     bool // Auto-fit content (default: false for fixed widths)

	// Conditional cell styling based on content
	// Map keys are matched case-insensitively against cell text
	// Matching cells will have their style overridden by the map value
	// Non-empty conditional values take precedence over row-level styling
	ConditionalStyles map[string]CellStyle
}

// ColumnDefinition defines properties for a table column
type ColumnDefinition struct {
	Title     string        // Column header title
	Width     int           // Optional: width in twips, 0 for auto
	Alignment CellAlignment // Optional: alignment for this column
	Bold      bool          // Make header bold
}

// CellStyle defines styling for table cells
type CellStyle struct {
	Bold       bool
	Italic     bool
	FontSize   int    // Font size in half-points (e.g., 20 = 10pt)
	FontColor  string // Hex color (e.g., "000000")
	Background string // Hex color for cell background
}

// InsertTable inserts a new table into the document
func (u *Updater) InsertTable(opts TableOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Validate options
	if err := validateTableOptions(opts); err != nil {
		return fmt.Errorf("invalid table options: %w", err)
	}

	// Set defaults
	opts = applyTableDefaults(opts)

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Generate table XML
	tableXML := generateTableXML(opts)

	// Insert table at the specified position
	updated, err := insertTableAtPosition(raw, tableXML, opts)
	if err != nil {
		return fmt.Errorf("insert table: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// validateTableOptions validates table creation options
func validateTableOptions(opts TableOptions) error {
	if len(opts.Columns) == 0 {
		return fmt.Errorf("at least one column is required")
	}

	// Check that all rows have the correct number of cells
	expectedCols := len(opts.Columns)
	for i, row := range opts.Rows {
		if len(row) != expectedCols {
			return fmt.Errorf("row %d has %d cells, expected %d", i, len(row), expectedCols)
		}
	}

	// Validate column widths if specified
	if len(opts.ColumnWidths) > 0 && len(opts.ColumnWidths) != expectedCols {
		return fmt.Errorf("column widths count (%d) must match columns count (%d)", len(opts.ColumnWidths), expectedCols)
	}

	return nil
}

// applyTableDefaults sets default values for unspecified options
func applyTableDefaults(opts TableOptions) TableOptions {
	if opts.BorderSize == 0 {
		opts.BorderSize = 4 // 0.5pt
	}
	if opts.BorderColor == "" {
		opts.BorderColor = "000000"
	}
	if opts.CellPadding == 0 {
		opts.CellPadding = 108 // 0.075 inch
	}
	if opts.BorderStyle == "" {
		opts.BorderStyle = BorderSingle
	}
	if opts.TableAlignment == "" {
		opts.TableAlignment = AlignLeft
	}
	if opts.HeaderAlignment == "" {
		opts.HeaderAlignment = CellAlignLeft
	}
	if opts.RowAlignment == "" {
		opts.RowAlignment = CellAlignLeft
	}
	if opts.VerticalAlign == "" {
		opts.VerticalAlign = VerticalAlignCenter
	}
	// Default: 100% width (spans between left and right margins)
	if opts.TableWidthType == "" {
		opts.TableWidthType = TableWidthPercentage
	}
	if opts.TableWidth == 0 {
		if opts.TableWidthType == TableWidthPercentage {
			opts.TableWidth = 5000 // 5000 = 100% in percentage mode
		}
		// For auto and fixed, 0 is a valid value
	}
	// Default row height rules
	if opts.HeaderHeightRule == "" {
		opts.HeaderHeightRule = RowHeightAuto
	}
	if opts.RowHeightRule == "" {
		opts.RowHeightRule = RowHeightAuto
	}

	return opts
}

// generateTableXML creates the complete XML for a table
func generateTableXML(opts TableOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString("<w:tbl>")

	// Table properties
	buf.WriteString("<w:tblPr>")

	// Table style
	if opts.TableStyle != "" {
		buf.WriteString(fmt.Sprintf(`<w:tblStyle w:val="%s"/>`, opts.TableStyle))
	}

	// Table width
	if opts.TableWidthType == TableWidthAuto {
		buf.WriteString(`<w:tblW w:w="0" w:type="auto"/>`)
	} else {
		buf.WriteString(fmt.Sprintf(`<w:tblW w:w="%d" w:type="%s"/>`, opts.TableWidth, opts.TableWidthType))
	}
	if opts.AutoFit {
		buf.WriteString(`<w:tblLayout w:type="autofit"/>`)
	} else {
		buf.WriteString(`<w:tblLayout w:type="fixed"/>`)
	}

	// Table alignment
	buf.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, opts.TableAlignment))

	// Table borders
	buf.WriteString(generateTableBorders(opts))

	// Cell margins (padding)
	buf.WriteString(fmt.Sprintf(`<w:tblCellMar>
		<w:top w:w="%d" w:type="dxa"/>
		<w:left w:w="%d" w:type="dxa"/>
		<w:bottom w:w="%d" w:type="dxa"/>
		<w:right w:w="%d" w:type="dxa"/>
	</w:tblCellMar>`, opts.CellPadding, opts.CellPadding, opts.CellPadding, opts.CellPadding))

	// Table look (for styled tables)
	buf.WriteString(`<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`)

	buf.WriteString("</w:tblPr>")

	// Table grid (column definitions)
	buf.WriteString(generateTableGrid(opts))

	// Header row
	buf.WriteString(generateHeaderRow(opts))

	// Data rows
	for i, rowData := range opts.Rows {
		isAlternate := (i % 2) == 1
		buf.WriteString(generateDataRow(opts, rowData, isAlternate))
	}

	buf.WriteString("</w:tbl>")

	return buf.Bytes()
}

// generateTableBorders creates border XML for the table
func generateTableBorders(opts TableOptions) string {
	if opts.BorderStyle == BorderNone {
		return `<w:tblBorders>
			<w:top w:val="none"/>
			<w:left w:val="none"/>
			<w:bottom w:val="none"/>
			<w:right w:val="none"/>
			<w:insideH w:val="none"/>
			<w:insideV w:val="none"/>
		</w:tblBorders>`
	}

	return fmt.Sprintf(`<w:tblBorders>
		<w:top w:val="%s" w:sz="%d" w:color="%s"/>
		<w:left w:val="%s" w:sz="%d" w:color="%s"/>
		<w:bottom w:val="%s" w:sz="%d" w:color="%s"/>
		<w:right w:val="%s" w:sz="%d" w:color="%s"/>
		<w:insideH w:val="%s" w:sz="%d" w:color="%s"/>
		<w:insideV w:val="%s" w:sz="%d" w:color="%s"/>
	</w:tblBorders>`,
		opts.BorderStyle, opts.BorderSize, opts.BorderColor,
		opts.BorderStyle, opts.BorderSize, opts.BorderColor,
		opts.BorderStyle, opts.BorderSize, opts.BorderColor,
		opts.BorderStyle, opts.BorderSize, opts.BorderColor,
		opts.BorderStyle, opts.BorderSize, opts.BorderColor,
		opts.BorderStyle, opts.BorderSize, opts.BorderColor)
}

// calculateProportionalColumnWidths calculates column widths based on content length
// Returns widths in twips, distributed proportionally to available space
func calculateProportionalColumnWidths(opts TableOptions, totalWidth int) []int {
	if len(opts.Columns) == 0 {
		return nil
	}

	// Calculate content length for each column
	// Each character roughly represents a certain width (default: 60 twips)
	contentLengths := make([]int, len(opts.Columns))
	totalContentLength := 0

	for i, col := range opts.Columns {
		// Start with header length
		length := len(col.Title)

		// Add length of longest cell in this column
		for _, row := range opts.Rows {
			if i < len(row) {
				cellLen := len(row[i])
				if cellLen > length {
					length = cellLen
				}
			}
		}

		contentLengths[i] = length
		totalContentLength += length
	}

	// Avoid division by zero
	if totalContentLength == 0 {
		totalContentLength = 1
	}

	// Distribute available width proportionally
	widths := make([]int, len(opts.Columns))
	distributedWidth := 0

	for i := 0; i < len(opts.Columns); i++ {
		// Calculate this column's proportion
		colWidth := (totalWidth * contentLengths[i]) / totalContentLength

		widths[i] = colWidth
		distributedWidth += colWidth
	}

	// Redistribute any rounding remainder to the last column
	if distributedWidth < totalWidth {
		widths[len(widths)-1] += totalWidth - distributedWidth
	}

	return widths
}

// generateTableGrid defines the table column structure
func generateTableGrid(opts TableOptions) string {
	var buf bytes.Buffer
	buf.WriteString("<w:tblGrid>")

	// Use specified widths or calculate
	if len(opts.ColumnWidths) > 0 {
		for _, width := range opts.ColumnWidths {
			buf.WriteString(fmt.Sprintf(`<w:gridCol w:w="%d"/>`, width))
		}
	} else if opts.ProportionalColumnWidths {
		// Calculate proportional widths based on content
		var totalWidth int

		if opts.TableWidthType == TableWidthPercentage {
			availableWidth := opts.AvailableWidth
			if availableWidth == 0 {
				availableWidth = 9360
			}
			totalWidth = (availableWidth * opts.TableWidth) / 5000
		} else if opts.TableWidthType == TableWidthFixed {
			totalWidth = opts.TableWidth
		} else {
			// For auto mode, use a reasonable default
			totalWidth = 11520
		}

		propWidths := calculateProportionalColumnWidths(opts, totalWidth)
		for _, width := range propWidths {
			buf.WriteString(fmt.Sprintf(`<w:gridCol w:w="%d"/>`, width))
		}
	} else {
		// Auto-calculate equal widths based on table width mode
		var colWidth int

		if opts.TableWidthType == TableWidthPercentage {
			// For percentage mode, distribute based on available page width
			// Grid columns must be in twips (absolute units) for proper sizing
			availableWidth := opts.AvailableWidth
			if availableWidth == 0 {
				// Default: Letter portrait (12240) - 1" margins (1440 each) = 9360 twips
				availableWidth = 9360
			}
			// Calculate proportion: how much of the available width this table takes
			// Percentage is 5000-based (5000 = 100%)
			tablePortionWidth := (availableWidth * opts.TableWidth) / 5000
			colWidth = tablePortionWidth / len(opts.Columns)
		} else if opts.TableWidthType == TableWidthFixed {
			// For fixed mode, distribute the specified fixed width
			colWidth = opts.TableWidth / len(opts.Columns)
		} else {
			// For auto mode, distribute a reasonable default width (8 inches = 11520 twips)
			colWidth = 11520 / len(opts.Columns)
		}

		for range opts.Columns {
			buf.WriteString(fmt.Sprintf(`<w:gridCol w:w="%d"/>`, colWidth))
		}
	}

	buf.WriteString("</w:tblGrid>")
	return buf.String()
}

// generateHeaderRow creates the table header row
func generateHeaderRow(opts TableOptions) string {
	var buf bytes.Buffer

	buf.WriteString("<w:tr>")

	// Row properties for header
	buf.WriteString("<w:trPr>")
	if opts.RepeatHeader {
		buf.WriteString("<w:tblHeader/>") // Repeat on each page
	}
	// Header row height
	if opts.HeaderRowHeight > 0 || opts.HeaderHeightRule != RowHeightAuto {
		height := opts.HeaderRowHeight
		if height == 0 {
			height = 360 // Default minimum if rule is specified but height is 0
		}
		buf.WriteString(fmt.Sprintf(`<w:trHeight w:val="%d" w:hRule="%s"/>`, height, opts.HeaderHeightRule))
	}
	buf.WriteString("</w:trPr>")

	// Header cells
	for i, col := range opts.Columns {
		alignment := opts.HeaderAlignment
		if col.Alignment != "" {
			alignment = col.Alignment
		}

		bold := opts.HeaderBold || col.Bold

		buf.WriteString(generateCell(
			col.Title,
			alignment,
			opts.VerticalAlign,
			opts.HeaderBackground,
			bold,
			false, // italic
			opts.HeaderStyle,
			opts.HeaderStyleName,
		))

		_ = i // unused but kept for potential future use
	}

	buf.WriteString("</w:tr>")
	return buf.String()
}

// resolveCellStyle determines the final cell style by merging row-level and conditional styles
func resolveCellStyle(cellContent string, rowStyle CellStyle, background string, conditionalStyles map[string]CellStyle) (CellStyle, string) {
	mergedStyle := rowStyle
	finalBackground := background

	// Check for conditional style match (case-insensitive)
	if conditionalStyles != nil {
		// Normalize cell content for comparison
		normalizedContent := strings.TrimSpace(cellContent)

		for key, condStyle := range conditionalStyles {
			if strings.EqualFold(normalizedContent, strings.TrimSpace(key)) {
				// Merge conditional style - conditional values override row defaults
				if condStyle.Background != "" {
					finalBackground = condStyle.Background
				}
				if condStyle.FontColor != "" {
					mergedStyle.FontColor = condStyle.FontColor
				}
				if condStyle.FontSize > 0 {
					mergedStyle.FontSize = condStyle.FontSize
				}
				// For booleans, use OR logic (either row or conditional can enable)
				mergedStyle.Bold = rowStyle.Bold || condStyle.Bold
				mergedStyle.Italic = rowStyle.Italic || condStyle.Italic
				break
			}
		}
	}

	return mergedStyle, finalBackground
}

// generateDataRow creates a table data row
func generateDataRow(opts TableOptions, rowData []string, isAlternate bool) string {
	var buf bytes.Buffer

	buf.WriteString("<w:tr>")

	// Row properties
	buf.WriteString("<w:trPr>")
	// Data row height
	if opts.RowHeight > 0 || opts.RowHeightRule != RowHeightAuto {
		height := opts.RowHeight
		if height == 0 {
			height = 360 // Default minimum if rule is specified but height is 0
		}
		buf.WriteString(fmt.Sprintf(`<w:trHeight w:val="%d" w:hRule="%s"/>`, height, opts.RowHeightRule))
	}
	buf.WriteString("</w:trPr>")

	// Determine background color
	background := opts.RowStyle.Background
	if isAlternate && opts.AlternateRowColor != "" {
		background = opts.AlternateRowColor
	}

	// Data cells
	for _, cellData := range rowData {
		// Resolve cell style (applying conditional formatting if applicable)
		cellStyle, cellBackground := resolveCellStyle(cellData, opts.RowStyle, background, opts.ConditionalStyles)

		buf.WriteString(generateCell(
			cellData,
			opts.RowAlignment,
			opts.VerticalAlign,
			cellBackground,
			cellStyle.Bold,
			cellStyle.Italic,
			cellStyle,
			opts.RowStyleName,
		))
	}

	buf.WriteString("</w:tr>")
	return buf.String()
}

// generateCell creates a single table cell
func generateCell(content string, align CellAlignment, vAlign VerticalAlignment, background string, bold, italic bool, style CellStyle, styleName string) string {
	var buf bytes.Buffer

	buf.WriteString("<w:tc>")

	// Cell properties
	buf.WriteString("<w:tcPr>")

	// Vertical alignment
	buf.WriteString(fmt.Sprintf(`<w:vAlign w:val="%s"/>`, vAlign))

	// Background color
	if background != "" {
		buf.WriteString(fmt.Sprintf(`<w:shd w:val="clear" w:color="auto" w:fill="%s"/>`, background))
	}

	buf.WriteString("</w:tcPr>")

	// Cell content (paragraph)
	buf.WriteString("<w:p>")

	// Paragraph properties (alignment and style)
	buf.WriteString("<w:pPr>")
	// Named paragraph style (if specified)
	if styleName != "" {
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, xmlEscape(styleName)))
	}
	buf.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, align))
	buf.WriteString("</w:pPr>")

	// Text run
	buf.WriteString("<w:r>")

	// Run properties (formatting)
	needsRPr := bold || italic || style.Bold || style.Italic || style.FontSize > 0 || style.FontColor != ""
	if needsRPr {
		buf.WriteString("<w:rPr>")
		if bold || style.Bold {
			buf.WriteString("<w:b/>")
		}
		if italic || style.Italic {
			buf.WriteString("<w:i/>")
		}
		if style.FontSize > 0 {
			buf.WriteString(fmt.Sprintf(`<w:sz w:val="%d"/>`, style.FontSize))
			buf.WriteString(fmt.Sprintf(`<w:szCs w:val="%d"/>`, style.FontSize))
		}
		if style.FontColor != "" {
			buf.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, style.FontColor))
		}
		buf.WriteString("</w:rPr>")
	}

	// Text content
	buf.WriteString("<w:t")
	if strings.HasPrefix(content, " ") || strings.HasSuffix(content, " ") {
		buf.WriteString(` xml:space="preserve"`)
	}
	buf.WriteString(">")
	buf.WriteString(xmlEscape(content))
	buf.WriteString("</w:t>")

	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	buf.WriteString("</w:tc>")

	return buf.String()
}

// insertTableAtPosition inserts the table XML at the specified position
func insertTableAtPosition(docXML, tableXML []byte, opts TableOptions) ([]byte, error) {
	// Handle caption if specified
	contentToInsert := tableXML
	if opts.Caption != nil {
		// Validate caption options
		if err := ValidateCaptionOptions(opts.Caption); err != nil {
			return nil, fmt.Errorf("invalid caption options: %w", err)
		}

		// Set caption type to Table if not already set
		if opts.Caption.Type == "" {
			opts.Caption.Type = CaptionTable
		}

		// Generate caption XML
		captionXML := generateCaptionXML(*opts.Caption)

		// Combine table and caption based on position
		contentToInsert = insertCaptionWithElement(docXML, captionXML, tableXML, opts.Caption.Position)
	}

	switch opts.Position {
	case PositionBeginning:
		return insertAtBodyStart(docXML, contentToInsert)
	case PositionEnd:
		return insertAtBodyEnd(docXML, contentToInsert)
	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionAfterText")
		}
		return insertAfterText(docXML, contentToInsert, opts.Anchor)
	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionBeforeText")
		}
		return insertBeforeText(docXML, contentToInsert, opts.Anchor)
	default:
		return nil, fmt.Errorf("invalid insert position")
	}
}

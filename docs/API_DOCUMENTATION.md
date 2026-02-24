# DOCX Chart Updater - API Documentation

## Go Fiber Backend Integration Guide

**Version**: 1.0.0
**Go Version**: 1.25.7+
**Import Path**: `github.com/falcomza/go-docx`

---

## Table of Contents

1. [Overview](#overview)
2. [Installation](#installation)
3. [Quick Start](#quick-start)
4. [Core Concepts](#core-concepts)
5. [API Reference](#api-reference)
6. [Fiber Integration](#fiber-integration)
7. [Error Handling](#error-handling)
8. [Best Practices](#best-practices)
9. [Performance Considerations](#performance-considerations)
10. [Type Reference](#type-reference)

---

## Overview

The DOCX Chart Updater is a Go library for programmatically manipulating Microsoft Word (DOCX) documents. It operates directly on the OpenXML format, enabling:

- **Chart Updates**: Modify chart data, categories, series, and titles
- **Chart Creation**: Insert new charts with embedded workbooks
- **Table Operations**: Insert styled tables with captions
- **Image Insertion**: Add images with proportional scaling and captions
- **Text Operations**: Insert paragraphs, find/replace text, read content
- **Document Structure**: Page breaks, section breaks, headers, footers
- **Hyperlinks**: Add clickable links to documents

### API Surface Area

| Category | Methods |
|----------|---------|
| **Lifecycle** | `New()`, `Save()`, `Cleanup()`, `TempDir()` |
| **Charts** | `UpdateChart()`, `InsertChart()` |
| **Tables** | `InsertTable()` |
| **Images** | `InsertImage()` |
| **Paragraphs** | `InsertParagraph()`, `InsertParagraphs()`, `AddHeading()`, `AddText()` |
| **Text Search/Replace** | `FindText()`, `ReplaceText()`, `ReplaceTextRegex()` |
| **Read Content** | `GetText()`, `GetParagraphText()`, `GetTableText()` |
| **Breaks** | `InsertPageBreak()`, `InsertSectionBreak()` |
| **Hyperlinks** | `InsertHyperlink()`, `InsertInternalLink()` |
| **Headers/Footers** | `SetHeader()`, `SetFooter()` |
| **Properties** | `SetCoreProperties()`, `SetAppProperties()`, `SetCustomProperties()`, `GetCoreProperties()` |
| **Bookmarks** | `CreateBookmark()`, `CreateBookmarkWithText()` |
| **Track Changes** | `InsertTrackedText()`, `DeleteTrackedText()` |
| **Deletion** | `DeleteParagraphs()`, `DeleteTable()`, `DeleteImage()`, `DeleteChart()` |
| **Table Update** | `UpdateTableCell()` |
| **Table Merge** | `MergeTableCellsHorizontal()`, `MergeTableCellsVertical()` |
| **Count** | `GetChartCount()`, `GetTableCount()`, `GetParagraphCount()`, `GetImageCount()` |
| **Chart Reading** | `GetChartData()` |
| **TOC** | `InsertTOC()`, `UpdateTOC()`, `GetTOCEntries()` |
| **Footnotes/Endnotes** | `InsertFootnote()`, `InsertEndnote()` |
| **Comments** | `InsertComment()`, `GetComments()` |
| **Styles** | `AddStyle()`, `AddStyles()` |
| **Page Numbers** | `SetPageNumber()` |
| **Watermarks** | `SetTextWatermark()` |

### Key Design Principles

- **1-based Indexing**: Chart indices start at 1, not 0
- **Strict Validation**: Fails fast on invalid input with descriptive errors
- **Resource Management**: Always use `defer updater.Cleanup()` for temp file cleanup
- **Thread-Safe**: Each `Updater` instance operates on isolated temp directories

---

## Installation

```bash
go get github.com/falcomza/go-docx@latest
```

### Import

```go
import godocx "github.com/falcomza/go-docx"
```

---

## Quick Start

### Basic Chart Update

```go
package main

import (
    "log"
    godocx "github.com/falcomza/go-docx"
)

func main() {
    // Create updater instance
    updater, err := godocx.New("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer updater.Cleanup()

    // Update first chart (1-based index)
    data := godocx.ChartData{
        Categories: []string{"Q1", "Q2", "Q3", "Q4"},
        Series: []godocx.SeriesData{
            {Name: "Revenue", Values: []float64{1000, 1500, 1200, 1800}},
            {Name: "Expenses", Values: []float64{800, 900, 850, 1000}},
        },
        ChartTitle:        "Quarterly Financial Results",
        CategoryAxisTitle: "Fiscal Quarter",
        ValueAxisTitle:    "Amount (USD)",
    }

    if err := updater.UpdateChart(1, data); err != nil {
        log.Fatal(err)
    }

    // Save to new file
    if err := updater.Save("output.docx"); err != nil {
        log.Fatal(err)
    }
}
```

---

## Core Concepts

### Updater Lifecycle

```go
// 1. Create - Opens DOCX, extracts to temp directory
updater, err := godocx.New("input.docx")

// 2. Modify - Perform operations on the document
updater.UpdateChart(1, data)
updater.InsertParagraph(opts)
updater.InsertTable(tableOpts)

// 3. Save - Writes modified DOCX to output path
updater.Save("output.docx")

// 4. Cleanup - Removes temporary files (use defer)
defer updater.Cleanup()
```

### Chart Indexing

Charts use **1-based indexing**:

```go
updater.UpdateChart(1, data)  // First chart
updater.UpdateChart(2, data)  // Second chart
// updater.UpdateChart(0, data)  // ERROR: index must be >= 1
```

### Workbook Resolution

Charts reference embedded Excel workbooks via relationships:

```
word/charts/chart1.xml
└── <c:externalData r:id="rId1"/>
    └── word/charts/_rels/chart1.xml.rels
        └── Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"
```

The library resolves these relationships automatically—no manual path handling required.

---

## API Reference

### Constructor

#### `New(docxPath string) (*Updater, error)`

Creates a new updater by extracting the DOCX to a temporary directory.

**Parameters:**
- `docxPath`: Path to the source DOCX file

**Returns:**
- `*Updater`: Updater instance for document manipulation
- `error`: File not found, invalid DOCX structure, or extraction failure

**Example:**
```go
updater, err := godocx.New("./templates/report.docx")
if err != nil {
    return fmt.Errorf("failed to open document: %w", err)
}
defer updater.Cleanup()
```

### Core Methods

#### `UpdateChart(chartIndex int, data ChartData) error`

Updates chart data and embedded Excel workbook for a specific chart.

**Parameters:**
- `chartIndex`: 1-based index of the chart to update
- `data`: Chart data containing categories, series, and optional titles

**Validation Rules:**
- Categories must not be empty
- At least one series required
- All series values must match categories length
- Series names cannot be empty/whitespace

**Example:**
```go
data := godocx.ChartData{
    Categories: []string{"Jan", "Feb", "Mar"},
    Series: []godocx.SeriesData{
        {Name: "Sales", Values: []float64{100, 200, 150}},
    },
}
if err := updater.UpdateChart(1, data); err != nil {
    return err
}
```

#### `Save(outputPath string) error`

Writes the modified document to the specified path.

**Parameters:**
- `outputPath`: Destination file path (creates parent directories if needed)

**Example:**
```go
if err := updater.Save("./output/final-report.docx"); err != nil {
    return fmt.Errorf("failed to save: %w", err)
}
```

#### `Cleanup() error`

Removes the temporary workspace. **Always call with `defer`**.

**Example:**
```go
updater, err := godocx.New("input.docx")
if err != nil {
    return err
}
defer updater.Cleanup() // Executed even on panic/early return
```

#### `TempDir() string`

Returns the temporary directory path for debugging/inspection.

### Paragraph Operations

#### `InsertParagraph(opts ParagraphOptions) error`

Inserts a single paragraph at the specified position.

**Options:**
```go
type ParagraphOptions struct {
    Text      string         // Required: paragraph content
    Style     ParagraphStyle // Default: Normal
    Position  InsertPosition // Default: End
    Anchor    string         // Required for PositionAfterText/BeforeText
    Bold      bool
    Italic    bool
    Underline bool
}
```

**Predefined Styles:**
- `StyleNormal`, `StyleHeading1`, `StyleHeading2`, `StyleHeading3`
- `StyleTitle`, `StyleSubtitle`, `StyleQuote`, `StyleListNumber`, `StyleListBullet`

**Example:**
```go
updater.InsertParagraph(godocx.ParagraphOptions{
    Text:     "Executive Summary",
    Style:    godocx.StyleHeading1,
    Position: godocx.PositionBeginning,
})
```

#### `InsertParagraphs(paragraphs []ParagraphOptions) error`

Batch inserts multiple paragraphs in sequence.

#### `AddHeading(level int, text string, position InsertPosition) error`

Convenience method for inserting heading paragraphs (level 1-3).

#### `AddText(text string, position InsertPosition) error`

Convenience method for inserting normal text paragraphs.

### Table Operations

#### `InsertTable(opts TableOptions) error`

Inserts a styled table with optional caption.

**Key Options:**
```go
type TableOptions struct {
    // Positioning
    Position  InsertPosition
    Anchor    string

    // Structure
    Columns      []ColumnDefinition
    ColumnWidths []int // nil for auto-calculated
    Rows         [][]string

    // Header styling
    HeaderStyle       CellStyle
    HeaderStyleName   string    // e.g., "Heading 1"
    RepeatHeader      bool
    HeaderBackground  string    // hex color
    HeaderBold        bool
    HeaderAlignment   CellAlignment

    // Row styling
    RowStyle          CellStyle
    AlternateRowColor string
    RowAlignment      CellAlignment
    VerticalAlign     VerticalAlignment

    // Table properties
    TableAlignment TableAlignment
    TableWidthType TableWidthType // auto/pct/dxa
    TableWidth     int            // 5000 = 100% in pct mode
    TableStyle     TableStyle

    // Borders
    BorderStyle BorderStyle
    BorderSize  int    // 4 = 0.5pt
    BorderColor string // hex

    // Spacing
    CellPadding int  // 108 = 0.075"
    AutoFit     bool

    // Caption (optional)
    Caption *CaptionOptions
}
```

**Example:**
```go
updater.InsertTable(godocx.TableOptions{
    Columns: []godocx.ColumnDefinition{
        {Title: "Metric", Width: 2000},
        {Title: "Value", Width: 1000},
    },
    Rows: [][]string{
        {"Revenue", "$1.2M"},
        {"Growth", "+15%"},
    },
    TableStyle:       godocx.TableStyleProfessional,
    HeaderBackground: "4472C4",
    HeaderBold:       true,
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionTable,
        Description: "Key performance indicators",
        AutoNumber:  true,
    },
})
```

### Image Operations

#### `InsertImage(opts ImageOptions) error`

Inserts an image with proportional scaling and optional caption.

**Options:**
```go
type ImageOptions struct {
    Path     string // Required: image file path
    Width    int    // Optional: pixels
    Height   int    // Optional: pixels
    AltText  string
    Position InsertPosition
    Anchor   string
    Caption  *CaptionOptions
}
```

**Supported Formats:** PNG, JPEG, GIF, BMP, TIFF

**Proportional Scaling:**
- Both specified → Uses exact dimensions
- Only width → Height calculated proportionally
- Only height → Width calculated proportionally
- Neither specified → Uses original image dimensions

**Example:**
```go
updater.InsertImage(godocx.ImageOptions{
    Path:     "./assets/logo.png",
    Width:    300,
    AltText:  "Company Logo",
    Position: godocx.PositionBeginning,
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionFigure,
        Description: "Company branding",
        AutoNumber:  true,
    },
})
```

### Text Search & Replace

#### `FindText(pattern string, opts FindOptions) ([]TextMatch, error)`

Searches for text with various options.

**Options:**
```go
type FindOptions struct {
    MatchCase    bool
    WholeWord    bool
    UseRegex     bool
    MaxResults   int // 0 = unlimited
    InParagraphs bool
    InTables     bool
    InHeaders    bool
    InFooters    bool
}
```

**Returns:**
```go
type TextMatch struct {
    Text      string
    Paragraph int
    Position  int
    Before    string
    After     string
}
```

**Example:**
```go
matches, err := updater.FindText("[TODO]", godocx.FindOptions{
    UseRegex:     true,
    InParagraphs: true,
    MaxResults:   10,
})
```

#### `ReplaceText(old, new string, opts ReplaceOptions) (int, error)`

Replaces text occurrences, returning count replaced.

#### `ReplaceTextRegex(pattern *regexp.Regexp, replacement string, opts ReplaceOptions) (int, error)`

Regex-based replacement.

### Reading Content

#### `GetText() (string, error)`

Extracts all visible text from the document body.

#### `GetParagraphText() ([]string, error)`

Returns text from each paragraph as a slice.

#### `GetTableText() ([][][]string, error)`

Returns table data as `tables[tableIndex][rowIndex][cellIndex]`.

### Break Operations

#### `InsertPageBreak(opts BreakOptions) error`

Inserts a page break.

#### `InsertSectionBreak(opts BreakOptions) error`

Inserts a section break.

**Section Types:**
- `SectionBreakNextPage` - New section on next page
- `SectionBreakContinuous` - Same page
- `SectionBreakEvenPage` - Next even page
- `SectionBreakOddPage` - Next odd page

### Hyperlink Operations

#### `InsertHyperlink(text, urlStr string, opts HyperlinkOptions) error`

Inserts an external hyperlink with styling options.

**Options:**
```go
type HyperlinkOptions struct {
    Position  InsertPosition
    Anchor    string
    Tooltip   string
    Style     ParagraphStyle
    Color     string // hex color, default: "0563C1" (Word blue)
    Underline bool
    ScreenTip string // accessibility text
}
```

**Example:**
```go
updater.InsertHyperlink("Visit our website", "https://example.com", godocx.HyperlinkOptions{
    Position:  godocx.PositionAfterText,
    Anchor:    "Contact Us",
    Color:     "0563C1",
    Underline: true,
    Tooltip:   "Opens in new tab",
})
```

#### `InsertInternalLink(text, bookmarkName string, opts HyperlinkOptions) error`

Inserts an internal link to a bookmark within the document.

### Header & Footer Operations

#### `SetHeader(content HeaderFooterContent, opts HeaderOptions) error`

Sets or creates a document header.

**Content Structure:**
```go
type HeaderFooterContent struct {
    LeftText          string
    CenterText        string
    RightText         string
    PageNumber        bool
    PageNumberFormat  string // e.g., "Page X of Y"
    Date              bool
    DateFormat        string // e.g., "MMMM d, yyyy"
}
```

**Header Options:**
```go
type HeaderOptions struct {
    Type             HeaderType // first, even, default
    DifferentFirst   bool
    DifferentOddEven bool
}
```

**Example:**
```go
updater.SetHeader(godocx.HeaderFooterContent{
    LeftText:   "Confidential Report",
    CenterText: "Q4 2024",
    RightText:  "Acme Corp",
    PageNumber: true,
    PageNumberFormat: "Page X of Y",
}, godocx.HeaderOptions{
    Type: godocx.HeaderDefault,
})
```

#### `SetFooter(content HeaderFooterContent, opts FooterOptions) error`

Sets or creates a document footer.

**Footer Types:**
- `FooterFirst` - First page only
- `FooterEven` - Even pages only
- `FooterDefault` - Odd pages (default)

**Example:**
```go
updater.SetFooter(godocx.HeaderFooterContent{
    CenterText: "Confidential - Do Not Distribute",
    PageNumber: true,
}, godocx.DefaultFooterOptions())
```

### Document Properties Operations

#### `SetCoreProperties(props CoreProperties) error`

Sets the core document properties (metadata).

**Properties:**
```go
type CoreProperties struct {
    Title          string    // Document title
    Subject        string    // Document subject
    Creator        string    // Author name
    Keywords       string    // Keywords (comma-separated)
    Description    string    // Description/comments
    Category       string    // Document category
    Created        time.Time // Creation date
    Modified       time.Time // Modification date
    LastModifiedBy string    // Last modifier name
    Revision       string    // Revision number
}
```

**Example:**
```go
updater.SetCoreProperties(godocx.CoreProperties{
    Title:       "Quarterly Financial Report",
    Subject:     "Q4 2024 Financials",
    Creator:     "John Doe",
    Keywords:    "finance, quarterly, report",
    Description: "Financial performance metrics for Q4 2024",
    Category:    "Reports",
})
```

#### `SetAppProperties(props AppProperties) error`

Sets application-specific document properties.

**Properties:**
```go
type AppProperties struct {
    Company     string // Company name
    Manager     string // Manager name
    Application string // Application name (typically Microsoft Word)
    AppVersion  string // Application version
}
```

**Example:**
```go
updater.SetAppProperties(godocx.AppProperties{
    Company:     "Acme Corporation",
    Manager:     "Jane Smith",
    Application: "Microsoft Word",
    AppVersion:  "16.0000",
})
```

#### `SetCustomProperties(properties []CustomProperty) error`

Sets custom document properties with typed values.

**Custom Property Structure:**
```go
type CustomProperty struct {
    Name  string      // Property name
    Value interface{} // Property value (string, int, float64, bool, or time.Time)
    Type  string      // Type (auto-inferred if empty)
}
```

**Supported Types:**
- `string` → "lpwstr"
- `int` → "i4"
- `float64` → "r8"
- `bool` → "bool"
- `time.Time` → "filetime"

**Example:**
```go
updater.SetCustomProperties([]godocx.CustomProperty{
    {Name: "ProjectCode", Value: "PRJ-2024-001"},
    {Name: "Budget", Value: 150000.50},
    {Name: "Approved", Value: true},
    {Name: "DueDate", Value: time.Date(2024, 12, 31, 0, 0, 0, 0, time.UTC)},
})
```

#### `GetCoreProperties() (*CoreProperties, error)`

Retrieves the current core document properties.

**Returns:**
- `*CoreProperties`: Current document properties
- `error`: Parse error or file not found

**Example:**
```go
props, err := updater.GetCoreProperties()
if err != nil {
    return err
}
fmt.Printf("Document Title: %s\n", props.Title)
fmt.Printf("Author: %s\n", props.Creator)
```

### Bookmark Operations

#### `CreateBookmark(name string, opts BookmarkOptions) error`

Creates an empty bookmark marker at the specified position.

**Options:**
```go
type BookmarkOptions struct {
    Position InsertPosition // Where to create bookmark
    Anchor   string         // Anchor text for relative positioning
    Style    ParagraphStyle // Style for bookmarked text
    Hidden   bool           // Invisible marker (default: true)
}
```

**Example:**
```go
// Create bookmark at document end
updater.CreateBookmark("section-start", godocx.BookmarkOptions{
    Position: godocx.PositionEnd,
    Hidden:   true,
})

// Create bookmark after specific text
updater.CreateBookmark("summary-section", godocx.BookmarkOptions{
    Position: godocx.PositionAfterText,
    Anchor:   "Executive Summary",
})
```

#### `CreateBookmarkWithText(name, text string, opts BookmarkOptions) error`

Creates a bookmark that wraps specific text content.

**Example:**
```go
updater.CreateBookmarkWithText("important-note", "Critical Information", godocx.BookmarkOptions{
    Position: godocx.PositionAfterText,
    Anchor:   "Introduction",
    Style:    godocx.StyleHeading2,
})
```

**Use Cases:**
- Navigation targets for internal hyperlinks
- Document structure markers
- Cross-reference anchors
- Table of contents generation

### Track Changes Operations

#### `InsertTrackedText(opts TrackedInsertOptions) error`

Inserts text with revision tracking enabled. The inserted text appears as a tracked insertion (typically shown with green underline in Word) and can be accepted or rejected by a reviewer.

**Options:**
```go
type TrackedInsertOptions struct {
    Text      string         // Required: text to insert
    Author    string         // Default: "Author"
    Date      time.Time      // Default: current time
    Position  InsertPosition // Where to insert
    Anchor    string         // Anchor text for relative positioning
    Style     ParagraphStyle // Style for the paragraph
    Bold      bool
    Italic    bool
    Underline bool
}
```

**Example:**
```go
updater.InsertTrackedText(godocx.TrackedInsertOptions{
    Text:      "This is newly added text.",
    Author:    "John Doe",
    Position:  godocx.PositionEnd,
    Bold:      true,
    Italic:    true,
})
```

#### `DeleteTrackedText(opts TrackedDeleteOptions) error`

Marks existing text as deleted with revision tracking. The text appears as struck-through red text and can be accepted or rejected by a reviewer.

**Options:**
```go
type TrackedDeleteOptions struct {
    Anchor string    // Required: text to mark as deleted
    Author string    // Default: "Author"
    Date   time.Time // Default: current time
}
```

**Example:**
```go
updater.DeleteTrackedText(godocx.TrackedDeleteOptions{
    Anchor: "Old text to remove",
    Author: "Jane Smith",
})
```

### Table Cell Operations

#### `UpdateTableCell(tableIndex, row, col int, value string) error`

Updates the text content of a specific cell in an existing table. All indices are 1-based.

**Note:** Nested tables (tables inside table cells) are not supported.

**Example:**
```go
// Update cell at table 1, row 2, column 3
updater.UpdateTableCell(1, 2, 3, "New cell content")
```

#### `MergeTableCellsHorizontal(tableIndex, row, startCol, endCol int) error`

Merges cells horizontally in a single row. All indices are 1-based.

**Example:**
```go
// Merge columns 2-4 in row 3 of table 1
updater.MergeTableCellsHorizontal(1, 3, 2, 4)
```

#### `MergeTableCellsVertical(tableIndex, startRow, endRow, col int) error`

Merges cells vertically in a single column. All indices are 1-based.

**Example:**
```go
// Merge rows 2-5 in column 1 of table 1
updater.MergeTableCellsVertical(1, 2, 5, 1)
```

### Content Deletion

#### `DeleteParagraphs(text string, opts DeleteOptions) (int, error)`

Removes paragraphs containing the specified text. Returns the number of paragraphs deleted.

**Options:**
```go
type DeleteOptions struct {
    MatchCase bool // Case-sensitive search
    WholeWord bool // Match whole words only
}
```

**Example:**
```go
count, err := updater.DeleteParagraphs("[DELETE]", godocx.DeleteOptions{
    MatchCase: false,
})
```

#### `DeleteTable(tableIndex int) error`

Removes a table by its 1-based index.

**Example:**
```go
updater.DeleteTable(2) // Delete second table
```

#### `DeleteImage(imageIndex int) error`

Removes an image by its 1-based index.

**Example:**
```go
updater.DeleteImage(1) // Delete first image
```

#### `DeleteChart(chartIndex int) error`

Removes a chart by its 1-based index.

**Example:**
```go
updater.DeleteChart(1) // Delete first chart
```

### Content Counting

#### `GetChartCount() (int, error)`

Returns the number of charts in the document.

#### `GetTableCount() (int, error)`

Returns the number of tables in the document.

#### `GetParagraphCount() (int, error)`

Returns the number of paragraphs in the document.

#### `GetImageCount() (int, error)`

Returns the number of images in the document.

### Chart Reading

#### `GetChartData(chartIndex int) (ChartData, error)`

Reads the categories, series names, and values from an existing chart.

**Example:**
```go
data, err := updater.GetChartData(1)
if err != nil {
    return err
}
fmt.Printf("Title: %s\n", data.ChartTitle)
fmt.Printf("Categories: %v\n", data.Categories)
for _, series := range data.Series {
    fmt.Printf("Series: %s = %v\n", series.Name, series.Values)
}
```

### Table of Contents Operations

#### `InsertTOC(opts TOCOptions) error`

Inserts a Table of Contents field into the document. The TOC uses Word field codes and populates when the document is opened in Word.

**Options:**
```go
type TOCOptions struct {
    Title         string // Default: "Table of Contents"
    OutlineLevels string // Default: "1-3" (include Heading1-3)
    Position      InsertPosition
    Anchor        string
}
```

**Example:**
```go
updater.InsertTOC(godocx.TOCOptions{
    Title:         "Table of Contents",
    OutlineLevels: "1-3",
    Position:      godocx.PositionBeginning,
})
```

#### `UpdateTOC() error`

Marks an existing TOC for update. When opened in Word, it will prompt to refresh.

#### `GetTOCEntries() ([]TOCEntry, error)`

Extracts TOC entries from the document.

**TOCEntry:**
```go
type TOCEntry struct {
    Level int
    Text  string
    Page  int
}
```

### Footnote and Endnote Operations

#### `InsertFootnote(opts FootnoteOptions) error`

Adds a footnote to the document.

**Options:**
```go
type FootnoteOptions struct {
    Text   string // Required: footnote content
    Anchor string // Required: text to attach footnote to
}
```

**Example:**
```go
updater.InsertFootnote(godocx.FootnoteOptions{
    Text:   "Source: Annual Report 2024",
    Anchor: "financial results",
})
```

#### `InsertEndnote(opts EndnoteOptions) error`

Adds an endnote to the document.

**Example:**
```go
updater.InsertEndnote(godocx.EndnoteOptions{
    Text:   "See Appendix A for details.",
    Anchor: "additional information",
})
```

### Comment Operations

#### `InsertComment(opts CommentOptions) error`

Adds a comment to the document attached to specific text.

**Options:**
```go
type CommentOptions struct {
    Text     string // Required: comment content
    Author   string // Default: "Author"
    Initials string // Default: first letter of Author
    Anchor   string // Required: text to attach comment to
}
```

**Example:**
```go
updater.InsertComment(godocx.CommentOptions{
    Text:     "Please verify this data.",
    Author:   "John Doe",
    Anchor:   "revenue figure",
})
```

#### `GetComments() ([]Comment, error)`

Retrieves all comments from the document.

**Comment:**
```go
type Comment struct {
    ID       int
    Author   string
    Initials string
    Date     string
    Text     string
}
```

### Custom Style Operations

#### `AddStyle(def StyleDefinition) error`

Adds a custom style to the document that can be used with paragraphs.

**StyleDefinition:**
```go
type StyleDefinition struct {
    ID          string         // Style ID (e.g., "CustomHeading")
    Name        string         // Display name
    Type        StyleType      // paragraph or character
    BasedOn     string         // Parent style
    NextStyle   string         // Next paragraph style
    FontFamily  string         // e.g., "Arial"
    FontSize    int            // half-points (24 = 12pt)
    Color       string         // hex color
    Bold        bool
    Italic      bool
    Underline   bool
    Alignment   ParagraphAlignment
    SpaceBefore int            // twips
    SpaceAfter  int            // twips
    LineSpacing int            // 240ths of line
    OutlineLevel int           // 1-9 for TOC
}
```

**Example:**
```go
updater.AddStyle(godocx.StyleDefinition{
    ID:          "CustomHeading",
    Name:        "Custom Heading",
    BasedOn:     "Heading1",
    FontFamily:  "Arial",
    FontSize:    36, // 18pt
    Color:       "FF0000",
    Bold:        true,
    Alignment:   godocx.AlignCenter,
    OutlineLevel: 1,
})
```

#### `AddStyles(defs []StyleDefinition) error`

Batch adds multiple custom styles.

### Page Number Operations

#### `SetPageNumber(opts PageNumberOptions) error`

Configures page numbering for the document.

**Options:**
```go
type PageNumberOptions struct {
    Start  int                // Starting page number (default: 1)
    Format PageNumberFormat   // decimal, upperRoman, lowerRoman, upperLetter, lowerLetter
}
```

**Example:**
```go
updater.SetPageNumber(godocx.PageNumberOptions{
    Start:  1,
    Format: godocx.PageNumDecimal,
})

// Roman numerals from page 5
updater.SetPageNumber(godocx.PageNumberOptions{
    Start:  5,
    Format: godocx.PageNumUpperRoman,
})
```

### Watermark Operations

#### `SetTextWatermark(opts WatermarkOptions) error`

Adds a diagonal text watermark to all pages.

**Options:**
```go
type WatermarkOptions struct {
    Text       string  // Required: watermark text
    FontFamily string  // Default: "Calibri"
    Color      string  // Default: "C0C0C0" (silver)
    Opacity    float64 // Default: 0.5 (0.0-1.0)
    Diagonal   bool    // Default: true (-45 degree rotation)
}
```

**Example:**
```go
updater.SetTextWatermark(godocx.WatermarkOptions{
    Text:     "CONFIDENTIAL",
    FontFamily: "Arial",
    Color:    "FF0000",
    Opacity:  0.3,
    Diagonal: true,
})
```

### Chart Creation

#### `InsertChart(opts ChartOptions) error`

Creates a new chart with embedded Excel workbook and inserts it into the document.

**Key Options:**
```go
type ChartOptions struct {
    // Positioning
    Position InsertPosition
    Anchor   string

    // Chart type
    ChartKind ChartKind // Column, Bar, Line, Pie, Area

    // Titles
    Title             string // Main chart title
    CategoryAxisTitle string // X-axis title
    ValueAxisTitle    string // Y-axis title

    // Data
    Categories []string     // Category labels
    Series     []SeriesData // Data series

    // Legend
    ShowLegend     bool   // Display legend (default: true)
    LegendPosition string // Position: "r", "l", "t", "b"

    // Dimensions (EMUs - English Metric Units)
    Width  int // Width in EMUs, 0 for default
    Height int // Height in EMUs, 0 for default

    // Caption
    Caption *CaptionOptions
}
```

**Chart Types:**
```go
const (
    ChartKindColumn ChartKind = "barChart"  // Vertical bars
    ChartKindBar    ChartKind = "barChart"  // Horizontal bars
    ChartKindLine   ChartKind = "lineChart" // Line chart
    ChartKindPie    ChartKind = "pieChart"  // Pie chart
    ChartKindArea   ChartKind = "areaChart" // Area chart
)
```

**Example:**
```go
updater.InsertChart(godocx.ChartOptions{
    Position:  godocx.PositionEnd,
    ChartKind: godocx.ChartKindColumn,
    Title:     "Sales Performance",
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []godocx.SeriesData{
        {Name: "2023", Values: []float64{100, 150, 120, 180}},
        {Name: "2024", Values: []float64{120, 170, 140, 200}},
    },
    ShowLegend:     true,
    LegendPosition: "r", // Right side
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionFigure,
        Description: "Quarterly sales comparison",
        AutoNumber:  true,
    },
})
```

**Differences from UpdateChart:**
- `InsertChart` creates a new chart from scratch
- `UpdateChart` modifies an existing chart in the template
- Use `InsertChart` when you need to add charts dynamically
- Use `UpdateChart` when working with pre-designed templates

---

## Fiber Integration

### Basic HTTP Handler

```go
package main

import (
    "github.com/gofiber/fiber/v2"
    godocx "github.com/falcomza/go-docx"
)

type UpdateChartRequest struct {
    ChartIndex int                    `json:"chartIndex"`
    Categories []string               `json:"categories"`
    Series     []SeriesDataRequest    `json:"series"`
    ChartTitle string                 `json:"chartTitle,omitempty"`
}

type SeriesDataRequest struct {
    Name   string    `json:"name"`
    Values []float64 `json:"values"`
}

func UpdateChartHandler(c *fiber.Ctx) error {
    var req UpdateChartRequest
    if err := c.BodyParser(&req); err != nil {
        return c.Status(400).JSON(fiber.Map{
            "error": "Invalid request body",
        })
    }

    // Get uploaded template
    file, err := c.FormFile("template")
    if err != nil {
        return c.Status(400).JSON(fiber.Map{
            "error": "Template file required",
        })
    }

    // Save temp file
    tempPath := fmt.Sprintf("./temp/%s", file.Filename)
    if err := c.SaveFile(file, tempPath); err != nil {
        return err
    }
    defer os.Remove(tempPath)

    // Process document
    updater, err := godocx.New(tempPath)
    if err != nil {
        return c.Status(500).JSON(fiber.Map{
            "error": "Failed to process document",
        })
    }
    defer updater.Cleanup()

    // Convert request to ChartData
    data := convertToChartData(req)

    // Update chart
    if err := updater.UpdateChart(req.ChartIndex, data); err != nil {
        return c.Status(500).JSON(fiber.Map{
            "error": err.Error(),
        })
    }

    // Save to output
    outputPath := fmt.Sprintf("./output/%s-updated.docx", file.Filename)
    if err := updater.Save(outputPath); err != nil {
        return err
    }

    // Return file
    return c.Download(outputPath)
}
```

### Streaming Response (Memory-Efficient)

```go
func StreamDocumentHandler(c *fiber.Ctx) error {
    // Process...
    outputPath := "./output/result.docx"

    // Stream file to client
    return c.SendFile(outputPath, true)
}
```

### Middleware Integration

```go
func DocxMiddleware() fiber.Handler {
    return func(c *fiber.Ctx) error {
        // Set up temp directory
        tempDir := filepath.Join(os.TempDir(), "docx-processing")
        os.MkdirAll(tempDir, 0755)

        // Store in context
        c.Locals("tempDir", tempDir)

        // Process request
        c.Next()

        // Cleanup (optional, depends on your strategy)
    }
}

app.Use(DocxMiddleware())
```

### Error Handling Middleware

```go
func ErrorHandler() fiber.Handler {
    return func(c *fiber.Ctx) error {
        err := c.Next()

        if err != nil {
            // Check for DocxError types
            var docxErr *godocx.DocxError
            if errors.As(err, &docxErr) {
                return c.Status(400).JSON(fiber.Map{
                    "code":    string(docxErr.Code),
                    "message": docxErr.Message,
                    "context": docxErr.Context,
                })
            }

            // Generic error
            return c.Status(500).JSON(fiber.Map{
                "error": err.Error(),
            })
        }

        return nil
    }
}
```

### Background Job Pattern

```go
// Queue for background processing
type DocxJob struct {
    ID        string
    Template  string
    Data      ChartData
    Status    string
    OutputURL string
}

var jobQueue = make(chan DocxJob, 100)

func ProcessJobsWorker() {
    for job := range jobQueue {
        updater, err := godocx.New(job.Template)
        if err != nil {
            job.Status = "failed"
            continue
        }

        updater.UpdateChart(1, job.Data)

        outputPath := fmt.Sprintf("./jobs/%s.docx", job.ID)
        updater.Save(outputPath)
        updater.Cleanup()

        job.Status = "completed"
        job.OutputURL = fmt.Sprintf("/download/%s", job.ID)
    }
}

// Start workers
for i := 0; i < 5; i++ {
    go ProcessJobsWorker()
}

// Handler to submit job
func SubmitJobHandler(c *fiber.Ctx) error {
    var req UpdateChartRequest
    c.BodyParser(&req)

    jobID := uuid.New().String()
    job := DocxJob{
        ID:       jobID,
        Template: saveUploadedFile(c),
        Data:     convertToChartData(req),
        Status:   "processing",
    }

    jobQueue <- job

    return c.JSON(fiber.Map{
        "jobId":  jobID,
        "status": "processing",
    })
}
```

---

## Error Handling

### Structured Error Types

```go
type DocxError struct {
    Code    ErrorCode
    Message string
    Err     error
    Context map[string]interface{}
}
```

### Error Codes

| Category | Code | Description |
|----------|------|-------------|
| Files | `INVALID_FILE` | Corrupted or invalid DOCX |
| Files | `FILE_NOT_FOUND` | Template missing |
| Charts | `CHART_NOT_FOUND` | Invalid chart index |
| Charts | `INVALID_CHART_DATA` | Data validation failed |
| Tables | `INVALID_TABLE_DATA` | Mismatched row/column counts |
| Images | `IMAGE_NOT_FOUND` | Image file missing |
| Images | `IMAGE_FORMAT` | Unsupported format |
| Text | `TEXT_NOT_FOUND` | Anchor text not found |
| Text | `INVALID_REGEX` | Pattern compilation failed |
| XML | `XML_PARSE` | Malformed XML |
| XML | `INVALID_XML` | Missing required structure |

### Error Handling Patterns

**Type Assertion for Specific Errors:**
```go
if err := updater.UpdateChart(1, data); err != nil {
    var docxErr *godocx.DocxError
    if errors.As(err, &docxErr) {
        switch docxErr.Code {
        case godocx.ErrCodeChartNotFound:
            return fmt.Errorf("chart %d does not exist", chartIndex)
        case godocx.ErrCodeInvalidChartData:
            return fmt.Errorf("data validation failed: %s", docxErr.Message)
        default:
            return err
        }
    }
    return err
}
```

**Context Extraction:**
```go
if docxErr, ok := err.(*godocx.DocxError); ok {
    if idx, exists := docxErr.Context["index"]; exists {
        log.Printf("Chart index: %v", idx)
    }
}
```

**Error Wrapping:**
```go
if err := updater.Save(outputPath); err != nil {
    return fmt.Errorf("failed to save report for client %d: %w", clientID, err)
}
```

---

## Best Practices

### 1. Always Defer Cleanup

```go
updater, err := godocx.New("template.docx")
if err != nil {
    return err
}
defer updater.Cleanup() // Guaranteed cleanup
```

### 2. Validate Input Early

```go
if len(data.Categories) == 0 {
    return godocx.NewInvalidChartDataError("categories required")
}
for i, series := range data.Series {
    if len(series.Values) != len(data.Categories) {
        return godocx.NewInvalidChartDataError(
            fmt.Sprintf("series %d length mismatch", i))
    }
}
```

### 3. Use Meaningful File Paths

```go
outputPath := filepath.Join(
    "./output",
    fmt.Sprintf("report_%s_%s.docx", clientID, time.Now().Format("20060102")),
)
```

### 4. Handle Concurrent Processing

```go
// Each updater gets isolated temp directory
wg := sync.WaitGroup{}
for _, template := range templates {
    wg.Add(1)
    go func(t string) {
        defer wg.Done()
        u, _ := godocx.New(t)
        defer u.Cleanup()
        // Process...
    }(template)
}
wg.Wait()
```

### 5. Preserve Original Templates

```go
// Copy template before modification
templateCopy := fmt.Sprintf("./temp/%s_copy.docx", uuid.New())
if err := copyFile(templatePath, templateCopy); err != nil {
    return err
}
defer os.Remove(templateCopy)

updater, err := godocx.New(templateCopy)
```

### 6. Set Reasonable Timeouts

```go
ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
defer cancel()

done := make(chan error, 1)
go func() {
    done <- processDocument(updater)
}()

select {
case err := <-done:
    return err
case <-ctx.Done():
    return fmt.Errorf("document processing timed out")
}
```

---

## Performance Considerations

### Memory Usage

- **Large Documents**: Extracted DOCX with many charts can use significant memory
- **Recommendation**: Process documents sequentially, not in parallel for single-instance use
- **Monitoring**: Track temp directory size with `updater.TempDir()`

### File I/O

- Each operation may read/write multiple XML files
- SSD storage significantly improves performance
- Network storage (NFS/SMB) may cause latency

### Optimization Tips

```go
// Batch operations to minimize file I/O
func ProcessDocumentBatch(updater *godocx.Updater, ops []Operation) error {
    // All modifications happen before Save()
    for _, op := range ops {
        if err := op.Apply(updater); err != nil {
            return err // Early exit, but Save() not yet called
        }
    }
    // Single Save() call
    return updater.Save(outputPath)
}
```

### Concurrency Pattern

```go
// Parallel processing of multiple documents
func ProcessParallel(inputs []string) error {
    sem := make(chan struct{}, runtime.NumCPU()) // Limit concurrency
    errChan := make(chan error, len(inputs))

    for _, input := range inputs {
        sem <- struct{}{} // Acquire
        go func(path string) {
            defer func() { <-sem }() // Release

            u, err := godocx.New(path)
            if err != nil {
                errChan <- err
                return
            }
            defer u.Cleanup()

            // Process...
            if err := u.Save(generateOutputPath(path)); err != nil {
                errChan <- err
            }
        }(input)
    }

    // Wait for completion
    for i := 0; i < cap(sem); i++ {
        sem <- struct{}{}
    }

    close(errChan)
    return <-errChan
}
```

---

## Type Reference

### Insert Position

```go
const (
    PositionBeginning InsertPosition = iota // Document start
    PositionEnd                             // Document end
    PositionAfterText                       // After anchor text
    PositionBeforeText                      // Before anchor text
)
```

### Table Styles

```go
const (
    TableStyleGrid         TableStyle = "TableGrid"
    TableStyleGridAccent1  TableStyle = "LightShading-Accent1"
    TableStylePlain        TableStyle = "TableNormal"
    TableStyleColorful     TableStyle = "ColorfulGrid-Accent1"
    TableStyleProfessional TableStyle = "LightGrid-Accent1"
)
```

### Cell Alignment

```go
const (
    CellAlignLeft   CellAlignment = "start"
    CellAlignCenter CellAlignment = "center"
    CellAlignRight  CellAlignment = "end"
)
```

### Caption Types

```go
const (
    CaptionFigure CaptionType = "Figure"
    CaptionTable  CaptionType = "Table"
)
```

### Section Break Types

```go
const (
    SectionBreakNextPage     SectionBreakType = "nextPage"
    SectionBreakContinuous   SectionBreakType = "continuous"
    SectionBreakEvenPage     SectionBreakType = "evenPage"
    SectionBreakOddPage      SectionBreakType = "oddPage"
)
```

### Header Types

```go
const (
    HeaderFirst    HeaderType = "first"    // First page header
    HeaderEven     HeaderType = "even"     // Even page header
    HeaderDefault  HeaderType = "default"  // Odd pages (default)
)
```

### Footer Types

```go
const (
    FooterFirst    FooterType = "first"    // First page footer
    FooterEven     FooterType = "even"     // Even page footer
    FooterDefault  FooterType = "default"  // Odd pages (default)
)
```

---

## Example: Complete Report Generation

```go
package main

import (
    "fmt"
    "log"
    godocx "github.com/falcomza/go-docx"
)

func GenerateReport(templatePath, outputPath string, data ReportData) error {
    // Initialize updater
    updater, err := godocx.New(templatePath)
    if err != nil {
        return fmt.Errorf("failed to load template: %w", err)
    }
    defer updater.Cleanup()

    // Replace placeholders
    updater.ReplaceText("{{COMPANY_NAME}}", data.Company, godocx.DefaultReplaceOptions())
    updater.ReplaceText("{{REPORT_DATE}}", data.Date.Format("2006-01-02"), godocx.DefaultReplaceOptions())

    // Update executive summary chart
    updater.UpdateChart(1, godocx.ChartData{
        Categories: data.Quarters,
        Series:     data.RevenueSeries,
        ChartTitle: "Revenue by Quarter",
    })

    // Insert KPI table
    updater.InsertTable(godocx.TableOptions{
        Columns: []godocx.ColumnDefinition{
            {Title: "Metric", Width: 2000, Bold: true},
            {Title: "Value", Width: 1500},
            {Title: "Change", Width: 1500},
        },
        Rows: data.KPIRows,
        HeaderStyle: godocx.CellStyle{
            Bold:       true,
            FontSize:   22,
            FontColor:  "FFFFFF",
            Background: "4472C4",
        },
        TableStyle: godocx.TableStyleProfessional,
        Caption: &godocx.CaptionOptions{
            Type:        godocx.CaptionTable,
            Description: "Key Performance Indicators",
            AutoNumber:  true,
        },
    })

    // Add chart for trend analysis
    if len(data.MonthlyTrends) > 0 {
        updater.InsertChart(godocx.ChartOptions{
            Position:   godocx.PositionEnd,
            ChartKind:  godocx.ChartKindColumn,
            Title:      "Monthly Trend Analysis",
            Categories: data.Months,
            Series:     data.TrendSeries,
            ShowLegend: true,
        })
    }

    // Insert logo
    if data.LogoPath != "" {
        updater.InsertImage(godocx.ImageOptions{
            Path:     data.LogoPath,
            Width:    200,
            Position: godocx.PositionBeginning,
        })
    }

    // Save output
    if err := updater.Save(outputPath); err != nil {
        return fmt.Errorf("failed to save report: %w", err)
    }

    return nil
}

type ReportData struct {
    Company       string
    Date          time.Time
    Quarters      []string
    RevenueSeries []godocx.SeriesData
    KPIRows       [][]string
    Months        []string
    TrendSeries   []godocx.SeriesData
    LogoPath      string
}
```

---

## Support & Contributing

- **Issues**: Report bugs at GitHub Issues
- **Documentation**: See `/docs` folder for additional guides
- **Examples**: Check `/examples` directory for code samples

---

## License

See LICENSE file for details.

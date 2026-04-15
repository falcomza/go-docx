# DOCX Updater

[![Go Version](https://img.shields.io/badge/Go-1.26+-00ADD8?style=flat&logo=go)](https://go.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful Go library for programmatically manipulating Microsoft Word (DOCX) documents. Update charts, insert tables, add paragraphs, generate captions, and more—all with a clean, idiomatic Go API.

## Features

🎯 **Document Content**
- **Paragraphs**: Styled text with headings (H1–H9), alignment, lists, and anchor positioning
- **Tables**: Formatted tables with custom styles, borders, row heights, and cell merging
- **Images**: Add images with automatic proportional sizing and flexible positioning
- **Hyperlinks & Bookmarks**: External URLs, internal links, and bookmark management
- **Lists**: Bullet and numbered lists with nesting support
- **Embedded Objects**: Embed Excel workbooks (and other OLE files) as interactive inline objects

📊 **Charts**
- **Chart Updates**: Modify existing chart data with automatic Excel workbook synchronization
- **Chart Insertion**: Create bar, column, line, pie, area, and scatter charts from scratch
- **Scatter Charts**: Full XValues support for true scatter/XY data
- **Multi-Chart Workflows**: Insert multiple charts programmatically
- **Read Chart Data**: Extract existing chart titles, categories, and series

📝 **Document Structure**
- **Table of Contents**: Generate automatic TOC using Word field codes, with update-on-open support
- **Table of Figures / Tables**: Generate caption-based lists for figure and table captions using Word field codes
- **Page & Section Breaks**: Control document flow with page and section breaks
- **Page Layout**: Configure page sizes, orientation, and margins per section
- **Headers & Footers**: Professional headers and footers with page numbering
- **Page Numbers**: Control starting page number and format (decimal, roman, letters)

✨ **Styles & Formatting**
- **Custom Styles**: Create paragraph and character styles with full formatting control
- **Text Watermarks**: Add diagonal or horizontal text watermarks via VML shapes
- **Auto-Captions**: Generate auto-numbered captions using Word's SEQ fields

📎 **Collaboration & Review**
- **Comments**: Add and read document comments with author and initials
- **Track Changes**: Insert text with revision tracking (insertions and deletions)
- **Footnotes & Endnotes**: Add scholarly footnotes and endnotes with reference markers

🔧 **Operations**
- **Text Find & Replace**: Search and replace with regex support
- **Read Operations**: Extract text from paragraphs, tables, headers, and footers
- **Delete Operations**: Remove paragraphs, tables, images, and charts by index
- **Update Operations**: Modify existing table cells
- **Count Operations**: Get counts of paragraphs, tables, images, and charts
- **Document Properties**: Full CRUD for core, app, and custom metadata (title, author, status, template, statistics, custom key-value pairs)

🛠️ **Advanced**
- **Blank Document Creation**: Create documents from scratch with `NewBlank()` — no template file needed
- **Template Upload**: Load templates from raw bytes with `NewFromBytes()` — ideal for web uploads and API payloads
- **`io.Reader`/`io.Writer` support**: In-memory document manipulation without disk I/O
- XML-based chart parsing using Go's `encoding/xml`
- Automatic Excel formula range adjustment
- Full OpenXML relationship and content type management
- Structured error types for better error handling
- Golden file tests for XML output verification

## Installation

```bash
go get github.com/falcomza/go-docx
```

## Quick Start

```go
package main

import (
    "log"
    godocx "github.com/falcomza/go-docx"
)

func main() {
    // Choose one constructor:
    u, err := godocx.New("template.docx")       // From existing file
    // u, err := godocx.NewBlank()               // From scratch (no template)
    // u, err := godocx.NewFromBytes(data)       // From raw bytes (upload/API)
    // u, err := godocx.NewFromReader(reader)    // From io.Reader
    if err != nil {
        log.Fatal(err)
    }
    defer u.Cleanup()

    // Update a chart
    chartData := godocx.ChartData{
        Categories: []string{"Q1", "Q2", "Q3", "Q4"},
        Series: []godocx.SeriesData{
            {Name: "Revenue", Values: []float64{100, 150, 120, 180}},
            {Name: "Costs", Values: []float64{80, 90, 85, 95}},
        },
    }
    u.UpdateChart(1, chartData)

    // Add a table
    u.InsertTable(godocx.TableOptions{
        Columns: []godocx.ColumnDefinition{
            {Title: "Product"},
            {Title: "Sales"},
            {Title: "Growth"},
        },
        Rows: [][]string{
            {"Product A", "$1.2M", "+15%"},
            {"Product B", "$980K", "+8%"},
        },
        TableStyle: godocx.TableStyleGridAccent1,
        Position:   godocx.PositionEnd,
        HeaderBold: true,
    })

    // Save result
    if err := u.Save("output.docx"); err != nil {
        log.Fatal(err)
    }
}
```

## Usage Examples

### Updating Chart Data

```go
u, _ := godocx.New("template.docx")
defer u.Cleanup()

data := godocx.ChartData{
    Categories: []string{"Jan", "Feb", "Mar", "Apr"},
    Series: []godocx.SeriesData{
        {Name: "Sales", Values: []float64{250, 300, 275, 350}},
    },
}

u.UpdateChart(1, data) // Update first chart (1-based index)
u.Save("updated.docx")
```

### Inserting Charts

Create charts from scratch, including scatter charts with custom X values:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// Column chart
u.InsertChart(godocx.ChartOptions{
    Title:      "Quarterly Revenue",
    ChartKind:  godocx.ChartKindColumn,
    Position:   godocx.PositionEnd,
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []godocx.SeriesOptions{
        {Name: "2025", Values: []float64{100, 120, 110, 130}},
        {Name: "2026", Values: []float64{110, 130, 125, 145}},
    },
})

// Scatter chart with custom X values
u.InsertChart(godocx.ChartOptions{
    Title:     "Correlation Analysis",
    ChartKind: godocx.ChartKindScatter,
    Position:  godocx.PositionEnd,
    ScatterChartOptions: &godocx.ScatterChartOptions{
        ScatterStyle: "smoothMarker",
    },
    Categories: []string{"Point 1", "Point 2", "Point 3"},
    Series: []godocx.SeriesOptions{
        {
            Name:    "Dataset A",
            Values:  []float64{10, 25, 40},
            XValues: []float64{1.5, 3.0, 4.5}, // Custom X values
        },
    },
})

u.Save("with_charts.docx")
```

**Supported chart types:** `ChartKindColumn` (vertical bars), `ChartKindBar` (horizontal bars), `ChartKindLine`, `ChartKindPie`, `ChartKindArea`, `ChartKindScatter`

> **Note:** `ChartKindColumn` and `ChartKindBar` are distinct constants — `Column` renders vertically (the default bar chart orientation) while `Bar` renders horizontally. Both emit `<c:barChart>` XML with the appropriate `barDir` attribute.

> **Note:** `ChartData` / `SeriesData` are used when *updating* existing charts (`UpdateChart`), while `ChartOptions` / `SeriesOptions` are used when *inserting* new charts (`InsertChart`).

### Creating Tables

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.InsertTable(godocx.TableOptions{
    Columns: []godocx.ColumnDefinition{
        {Title: "Product", Width: 2000, Bold: true},
        {Title: "Q1", Width: 1200},
        {Title: "Q2", Width: 1200},
    },
    Rows: [][]string{
        {"Product A", "$120K", "$135K"},
        {"Product B", "$98K", "$105K"},
    },
    TableStyle: godocx.TableStyleGridAccent1,
    Position:   godocx.PositionEnd,
    HeaderBold: true,
    RowHeight:  280,
})

// Update an existing cell
u.UpdateTableCell(1, 2, 3, "$140K") // table 1, row 2, col 3

// Merge cells horizontally (columns 1-3 in row 1)
u.MergeTableCellsHorizontal(1, 1, 1, 3)

// Merge cells vertically (rows 1-3 in column 1)
u.MergeTableCellsVertical(1, 1, 3, 1)

u.Save("with_table.docx")
```

### Adding Paragraphs

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.AddHeading(1, "Executive Summary", godocx.PositionEnd)
u.AddHeading(2, "Background", godocx.PositionEnd)
u.AddHeading(3, "Scope", godocx.PositionEnd)
u.AddHeading(4, "Methodology – Data Collection", godocx.PositionEnd) // H1–H9 supported

u.AddText("This quarter showed strong growth.", godocx.PositionEnd)

u.InsertParagraph(godocx.ParagraphOptions{
    Text:      "Important: Review required",
    Bold:      true,
    Italic:    true,
    Alignment: godocx.ParagraphAlignCenter,
    Position:  godocx.PositionEnd,
})

// Anchor-based positioning (works across split Word runs)
u.InsertParagraph(godocx.ParagraphOptions{
    Text:     "Follow-up details",
    Position: godocx.PositionAfterText,
    Anchor:   "Executive Summary",
})

// Newlines and tabs are emitted as <w:br/> and <w:tab/>
u.InsertParagraph(godocx.ParagraphOptions{
    Text:     "Line 1\nLine 2\tTabbed",
    Position: godocx.PositionEnd,
})

u.Save("with_paragraphs.docx")
```

### Table of Contents

Generate an automatic Table of Contents using Word field codes:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// Insert TOC at the beginning
u.InsertTOC(godocx.TOCOptions{
    Title:         "Table of Contents",
    OutlineLevels: "1-3",
    Position:      godocx.PositionBeginning,
})

// Add headings (these will appear in the TOC)
u.AddHeading(1, "Chapter 1: Introduction", godocx.PositionEnd)
u.AddText("Introduction content...", godocx.PositionEnd)
u.AddHeading(2, "1.1 Background", godocx.PositionEnd)
u.AddText("Background content...", godocx.PositionEnd)

// Mark TOC for update (Word recalculates on open)
u.UpdateTOC()

// Read existing TOC entries
entries, _ := u.GetTOCEntries()
for _, entry := range entries {
    fmt.Printf("Level %d: %s\n", entry.Level, entry.Text)
}

u.Save("with_toc.docx")
```

### Table of Figures and Table of Tables

Generate caption-based lists for existing `Figure` and `Table` captions:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.InsertTableOfFigures(godocx.DefaultTableOfFiguresOptions())
u.InsertTableOfTables(godocx.DefaultTableOfTablesOptions())

// Word populates the lists when fields are updated on open
u.UpdateTOC()
u.Save("with_caption_lists.docx")
```

### Custom Styles

Create and apply custom paragraph and character styles:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.AddStyles([]godocx.StyleDefinition{
    {
        ID:         "DocTitle",
        Name:       "Document Title",
        Type:       godocx.StyleTypeParagraph,
        BasedOn:    "Normal",
        FontFamily: "Calibri",
        FontSize:   56, // half-points (56 = 28pt)
        Color:      "1F4E79",
        Bold:       true,
        Alignment:  godocx.ParagraphAlignCenter,
        SpaceAfter: 240,
    },
    {
        ID:     "Highlight",
        Name:   "Strong Emphasis",
        Type:   godocx.StyleTypeCharacter,
        Bold:   true,
        Italic: true,
        Color:  "C00000",
    },
})

// Use the custom style
u.InsertParagraph(godocx.ParagraphOptions{
    Text:     "My Document Title",
    Style:    "DocTitle",
    Position: godocx.PositionBeginning,
})

u.Save("with_styles.docx")
```

### Watermarks

Add text watermarks to documents:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.SetTextWatermark(godocx.WatermarkOptions{
    Text:       "CONFIDENTIAL",
    FontFamily: "Calibri",
    Color:      "C0C0C0",
    Opacity:    0.3,
    Diagonal:   true, // 315-degree rotation
})

u.Save("watermarked.docx")
```

### Page Numbers

Control page numbering format and starting number:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.SetPageNumber(godocx.PageNumberOptions{
    Start:  1,
    Format: godocx.PageNumDecimal, // also: PageNumUpperRoman, PageNumLowerRoman, etc.
})

u.Save("with_page_numbers.docx")
```

### Footnotes and Endnotes

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.AddText("The experiment showed significant results.", godocx.PositionEnd)

// Insert footnote at anchor text
u.InsertFootnote(godocx.FootnoteOptions{
    Text:   "Based on data collected in Q3 2026.",
    Anchor: "significant results",
})

// Insert endnote at anchor text
u.InsertEndnote(godocx.EndnoteOptions{
    Text:   "See full methodology in Appendix A.",
    Anchor: "experiment",
})

u.Save("with_notes.docx")
```

### Comments

Add and read document comments:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.AddText("Revenue grew 15% this quarter.", godocx.PositionEnd)

u.InsertComment(godocx.CommentOptions{
    Text:     "Please verify this figure with accounting.",
    Author:   "Jane Reviewer",
    Initials: "JR",
    Anchor:   "grew 15%",
})

// Read existing comments
comments, _ := u.GetComments()
for _, c := range comments {
    fmt.Printf("%s: %s\n", c.Author, c.Text)
}

u.Save("with_comments.docx")
```

### Track Changes

Insert and delete text with revision tracking:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// Insert tracked text (green underline in Word)
u.InsertTrackedText(godocx.TrackedInsertOptions{
    Text:     "This paragraph was added during review.",
    Author:   "Jane Reviewer",
    Date:     time.Now(),
    Position: godocx.PositionEnd,
    Bold:     true,
})

// Mark existing text as deleted (red strikethrough in Word)
u.DeleteTrackedText(godocx.TrackedDeleteOptions{
    Anchor: "paragraph to be removed",
    Author: "Jane Reviewer",
    Date:   time.Now(),
})

u.Save("with_tracked_changes.docx")
```

### Delete Operations

Remove content from documents:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// Delete paragraphs containing specific text
count, _ := u.DeleteParagraphs("draft", godocx.DeleteOptions{MatchCase: false})
fmt.Printf("Deleted %d paragraphs\n", count)

// Delete by index (1-based)
u.DeleteTable(2)  // Remove 2nd table
u.DeleteImage(1)  // Remove 1st image
u.DeleteChart(1)  // Remove 1st chart

// Count operations
tableCount, _ := u.GetTableCount()
paraCount, _ := u.GetParagraphCount()
imageCount, _ := u.GetImageCount()
chartCount, _ := u.GetChartCount()

u.Save("cleaned.docx")
```

### io.Reader / io.Writer Support

Work with documents in memory without disk I/O:

```go
import (
    "bytes"
    "os"
    godocx "github.com/falcomza/go-docx"
)

// Open from io.Reader (e.g., HTTP upload, S3 object, etc.)
file, _ := os.Open("template.docx")
defer file.Close()

u, _ := godocx.NewFromReader(file)
defer u.Cleanup()

u.AddText("Added via io.Reader", godocx.PositionEnd)

// Save to io.Writer (e.g., HTTP response, S3 upload, etc.)
var buf bytes.Buffer
u.SaveToWriter(&buf)

// buf.Bytes() contains the complete DOCX file
os.WriteFile("output.docx", buf.Bytes(), 0o644)
```

### Inserting Images

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// Width only — height calculated proportionally
u.InsertImage(godocx.ImageOptions{
    Path:     "images/logo.png",
    Width:    400,
    AltText:  "Company Logo",
    Position: godocx.PositionEnd,
})

// With auto-numbered caption
u.InsertImage(godocx.ImageOptions{
    Path:     "images/chart.png",
    Width:    500,
    Position: godocx.PositionEnd,
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionFigure,
        Description: "Q1 Sales Performance",
        AutoNumber:  true,
    },
})

u.Save("with_images.docx")
```

**Proportional sizing:** Specify only `Width` (height auto-calculated), only `Height` (width auto-calculated), both (used as-is), or neither (actual image dimensions). Supported formats: PNG, JPEG, GIF, BMP, TIFF.

### Embedded Object Insertion

Embed an Excel workbook (or other OLE-supported file) as a double-clickable object:

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// Embed from file path (uses built-in Excel icon)
u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
    FilePath: "data/report.xlsx",
    Position: godocx.PositionEnd,
})

// Embed from bytes with custom display size
u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
    FileBytes: xlsxBytes,
    FileName:  "quarterly_report.xlsx",
    Width:     120,
    Height:    90,
    Position:  godocx.PositionAfterText,
    Anchor:    "See attached data:",
})

// Embed with a custom icon image
u.InsertEmbeddedObject(godocx.EmbeddedObjectOptions{
    FilePath:  "data/report.xlsx",
    IconPath:  "assets/excel_icon.png",
    Position:  godocx.PositionEnd,
})

u.Save("with_embedded.docx")
```

### Page and Section Breaks

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.InsertPageBreak(godocx.BreakOptions{
    Position: godocx.PositionEnd,
})

u.InsertSectionBreak(godocx.BreakOptions{
    Position:    godocx.PositionEnd,
    SectionType: godocx.SectionBreakNextPage,
    PageLayout:  godocx.PageLayoutA3Landscape(),
})

u.Save("with_breaks.docx")
```

**Section break types:** `SectionBreakNextPage`, `SectionBreakContinuous`, `SectionBreakEvenPage`, `SectionBreakOddPage`

**Layout helpers:** `PageLayoutLetterPortrait()`, `PageLayoutLetterLandscape()`, `PageLayoutA4Portrait()`, `PageLayoutA4Landscape()`, `PageLayoutA3Portrait()`, `PageLayoutA3Landscape()`, `PageLayoutLegalPortrait()`

### Hyperlinks and Bookmarks

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

// External hyperlink
u.InsertHyperlink("Visit GitHub", "https://github.com/falcomza/go-docx", godocx.HyperlinkOptions{
    Position:  godocx.PositionEnd,
    Color:     "0563C1",
    Underline: true,
    Tooltip:   "Open repository",
})

// Create bookmark with heading text
u.CreateBookmarkWithText("summary", "Executive Summary", godocx.BookmarkOptions{
    Position: godocx.PositionEnd,
    Style:    godocx.StyleHeading1,
})

// Internal link to bookmark
u.InsertInternalLink("Go to Summary", "summary", godocx.HyperlinkOptions{
    Position:  godocx.PositionBeginning,
    Color:     "0563C1",
    Underline: true,
})

u.Save("with_links.docx")
```

### Headers and Footers

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.SetHeader(godocx.HeaderFooterContent{
    LeftText:   "Company Name",
    CenterText: "Confidential Report",
    RightText:  "Feb 2026",
}, godocx.DefaultHeaderOptions())

u.SetFooter(godocx.HeaderFooterContent{
    CenterText:       "Page ",
    PageNumber:       true,
    PageNumberFormat: "X of Y",
}, godocx.DefaultFooterOptions())

u.Save("with_headers_footers.docx")
```

### Text Find & Replace

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

opts := godocx.DefaultReplaceOptions()
count, _ := u.ReplaceText("{{name}}", "John Doe", opts)

// Regex replacement
pattern := regexp.MustCompile(`\d{3}-\d{3}-\d{4}`)
count, _ = u.ReplaceTextRegex(pattern, "[REDACTED]", opts)

u.Save("replaced.docx")
```

### Read Operations

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

text, _ := u.GetText()                // All document text
paragraphs, _ := u.GetParagraphText()  // Text by paragraphs
tables, _ := u.GetTableText()          // Text from tables

// Find text with context
opts := godocx.DefaultFindOptions()
matches, _ := u.FindText("TODO:", opts)
for _, m := range matches {
    fmt.Printf("Paragraph %d: %s\n", m.Paragraph, m.Text)
}
```

### Creating Documents from Scratch

Create a blank document without any template file:

```go
u, _ := godocx.NewBlank()
defer u.Cleanup()

u.AddHeading(1, "Report Title", godocx.PositionEnd)
u.AddText("Created entirely from scratch.", godocx.PositionEnd)
u.Save("from_scratch.docx")
```

Load a template from raw bytes (e.g., HTTP upload or database):

```go
data, _ := io.ReadAll(uploadedFile)
u, _ := godocx.NewFromBytes(data)
defer u.Cleanup()

u.AddText("Added to uploaded template.", godocx.PositionEnd)
u.Save("from_upload.docx")
```

### Document Properties

Set and read core, application, and custom properties — matching Word's File → Info → Properties dialog:

```go
u, _ := godocx.NewBlank()
defer u.Cleanup()

// Core properties (Summary tab)
u.SetCoreProperties(godocx.CoreProperties{
    Title:         "Q4 Report",
    Creator:       "Finance Dept",
    Subject:       "Quarterly Financials",
    Keywords:      "finance, quarterly, report",
    ContentStatus: "Draft", // "Draft", "Final", "Reviewed", etc.
})

// Application properties (Statistics tab + template assignment)
u.SetAppProperties(godocx.AppProperties{
    Company:              "ACME Corp",
    Manager:              "Jane Smith",
    Template:             "Corporate_Report.dotm",
    HyperlinkBase:        "https://docs.acme.com",
    Pages:                25,
    Words:                7500,
    Characters:           42000,
    CharactersWithSpaces: 49500,
    Lines:                350,
    Paragraphs:           80,
    TotalTime:            120,
})

// Custom properties (Custom tab) — typed values
u.SetCustomProperties([]godocx.CustomProperty{
    {Name: "Department", Value: "Engineering"},
    {Name: "ProjectCode", Value: "PRJ-2026"},
    {Name: "Budget", Value: 150000.50},
    {Name: "Approved", Value: true},
    {Name: "Deadline", Value: time.Date(2026, 12, 31, 0, 0, 0, 0, time.UTC)},
})

// Read properties back
coreProps, _ := u.GetCoreProperties()
fmt.Println(coreProps.Title, coreProps.ContentStatus)

appProps, _ := u.GetAppProperties()
fmt.Println(appProps.Company, appProps.Template, appProps.Pages)

customProps, _ := u.GetCustomProperties()
for _, p := range customProps {
    fmt.Printf("%s = %v\n", p.Name, p.Value)
}

u.Save("with_properties.docx")
```

### Lists

```go
u, _ := godocx.New("document.docx")
defer u.Cleanup()

u.AddBulletList([]string{
    "First item",
    "Second item",
    "Third item",
}, 0, godocx.PositionEnd)

u.AddNumberedList([]string{
    "Step 1: Planning",
    "Step 2: Development",
    "Step 3: Testing",
}, 0, godocx.PositionEnd)

u.Save("with_lists.docx")
```

## API Overview

### Core Operations
| Method | Description |
|--------|-------------|
| `New(filepath string)` | Open DOCX file from disk |
| `NewBlank()` | Create blank document from scratch (no template needed) |
| `NewFromBytes(data []byte)` | Create from raw bytes (upload/API/database) |
| `NewFromReader(r io.Reader)` | Open DOCX from any `io.Reader` |
| `Save(outputPath string)` | Save document to disk |
| `SaveToWriter(w io.Writer)` | Save document to any `io.Writer` |
| `Cleanup()` | Clean up temporary files |

### Paragraph Operations
| Method | Description |
|--------|-------------|
| `InsertParagraph(opts ParagraphOptions)` | Insert styled paragraph |
| `InsertParagraphs(paragraphs []ParagraphOptions)` | Insert multiple paragraphs |
| `AddHeading(level, text, position)` | Insert heading at level 1–9 (matches Word's built-in Heading 1 – Heading 9 styles) |
| `AddText(text, position)` | Insert normal text |
| `AddBulletItem(text, level, position)` | Insert bullet item |
| `AddBulletList(items, level, position)` | Insert bullet list |
| `AddNumberedItem(text, level, position)` | Insert numbered item |
| `AddNumberedList(items, level, position)` | Insert numbered list |

### Table Operations
| Method | Description |
|--------|-------------|
| `InsertTable(opts TableOptions)` | Insert formatted table |
| `UpdateTableCell(table, row, col, value)` | Modify existing cell |
| `MergeTableCellsHorizontal(table, row, startCol, endCol)` | Merge cells across columns |
| `MergeTableCellsVertical(table, startRow, endRow, col)` | Merge cells across rows |

### Chart Operations
| Method | Description |
|--------|-------------|
| `InsertChart(opts ChartOptions)` | Create new chart |
| `UpdateChart(index, data)` | Update existing chart data |
| `GetChartCount()` | Count charts in document |
| `GetChartData(chartIndex)` | Read chart title, categories, and series |

### Table of Contents
| Method | Description |
|--------|-------------|
| `InsertTOC(opts TOCOptions)` | Insert TOC field |
| `InsertTableOfFigures(opts CaptionListOptions)` | Insert caption-based list for figure captions |
| `InsertTableOfTables(opts CaptionListOptions)` | Insert caption-based list for table captions |
| `UpdateTOC()` | Mark TOC for recalculation on open |
| `GetTOCEntries()` | Parse existing TOC entries |

### Styles
| Method | Description |
|--------|-------------|
| `AddStyle(def StyleDefinition)` | Add single custom style |
| `AddStyles(defs []StyleDefinition)` | Add multiple custom styles |

### Comments
| Method | Description |
|--------|-------------|
| `InsertComment(opts CommentOptions)` | Add comment at anchor text |
| `GetComments()` | Read all document comments |

### Track Changes
| Method | Description |
|--------|-------------|
| `InsertTrackedText(opts TrackedInsertOptions)` | Insert text with revision tracking |
| `DeleteTrackedText(opts TrackedDeleteOptions)` | Mark text as tracked deletion |

### Footnotes & Endnotes
| Method | Description |
|--------|-------------|
| `InsertFootnote(opts FootnoteOptions)` | Add footnote at anchor text |
| `InsertEndnote(opts EndnoteOptions)` | Add endnote at anchor text |

### Image Operations
| Method | Description |
|--------|-------------|
| `InsertImage(opts ImageOptions)` | Insert image with proportional sizing |

### Embedded Object Operations
| Method | Description |
|--------|-------------|
| `InsertEmbeddedObject(opts EmbeddedObjectOptions)` | Embed an OLE object (e.g., Excel workbook) as a double-clickable inline icon |

### Hyperlink & Bookmark Operations
| Method | Description |
|--------|-------------|
| `InsertHyperlink(text, url, opts)` | Insert external hyperlink |
| `InsertInternalLink(text, bookmark, opts)` | Insert internal link |
| `CreateBookmark(name, opts)` | Create empty bookmark |
| `CreateBookmarkWithText(name, text, opts)` | Create bookmark with content |
| `WrapTextInBookmark(name, anchorText)` | Wrap existing text in bookmark |

### Text Operations
| Method | Description |
|--------|-------------|
| `ReplaceText(old, new, opts)` | Replace text occurrences |
| `ReplaceTextRegex(pattern, replacement, opts)` | Replace using regex |
| `GetText()` | Extract all document text |
| `GetParagraphText()` | Extract text by paragraphs |
| `GetTableText()` | Extract text from tables |
| `FindText(pattern, opts)` | Find text with context |

### Delete Operations
| Method | Description |
|--------|-------------|
| `DeleteParagraphs(text, opts)` | Delete paragraphs matching text |
| `DeleteTable(index)` | Delete table by index |
| `DeleteImage(index)` | Delete image by index |
| `DeleteChart(index)` | Delete chart by index |

### Count Operations
| Method | Description |
|--------|-------------|
| `GetTableCount()` | Count tables in document |
| `GetParagraphCount()` | Count paragraphs |
| `GetImageCount()` | Count images |
| `GetChartCount()` | Count charts |

### Page Layout & Formatting
| Method | Description |
|--------|-------------|
| `SetPageNumber(opts PageNumberOptions)` | Set page number start and format |
| `SetTextWatermark(opts WatermarkOptions)` | Add text watermark |
| `SetPageLayout(opts PageLayoutOptions)` | Set page size and orientation |
| `InsertPageBreak(opts BreakOptions)` | Insert page break |
| `InsertSectionBreak(opts BreakOptions)` | Insert section break |

### Header & Footer Operations
| Method | Description |
|--------|-------------|
| `SetHeader(content, opts)` | Create/update header |
| `SetFooter(content, opts)` | Create/update footer |

### Properties Operations
| Method | Description |
|--------|-------------|
| `SetCoreProperties(props)` | Set core metadata (Title, Author, ContentStatus, etc.) |
| `GetCoreProperties()` | Read core metadata |
| `SetAppProperties(props)` | Set app metadata (Company, Template, statistics, etc.) |
| `GetAppProperties()` | Read app metadata |
| `SetCustomProperties(properties)` | Set custom key-value metadata |
| `GetCustomProperties()` | Read custom key-value metadata with preserved types |

### Caption Operations
| Method | Description |
|--------|-------------|
| `AddCaption(opts CaptionOptions)` | Insert auto-numbered caption |

## Project Structure

```
.
├── chart_updater.go     # Main Updater API, New/Save/io.Reader/io.Writer
├── chart.go             # Chart insertion (column, bar, line, pie, area, scatter)
├── chart_xml.go         # XML manipulation for charts
├── chart_read.go        # Read existing chart data
├── chart_extended.go    # Extended chart types and options
├── excel_handler.go     # Embedded workbook updates
├── table.go             # Table insertion with styles
├── table_update.go      # Update existing table cells
├── merge.go             # Table cell merging (horizontal/vertical)
├── paragraph.go         # Paragraph and text insertion
├── image.go             # Image insertion with proportional sizing
├── embed.go             # OLE embedded object insertion
├── toc.go               # Table of Contents generation
├── styles.go            # Custom style definitions
├── watermark.go         # Text watermarks via VML
├── pagenumber.go        # Page number control
├── footnote.go          # Footnotes and endnotes
├── comment.go           # Document comments
├── trackchanges.go      # Revision tracking (insertions/deletions)
├── delete.go            # Delete operations and count queries
├── bookmark.go          # Bookmark management
├── hyperlink.go         # Hyperlinks (external and internal)
├── headerfooter.go      # Headers and footers
├── breaks.go            # Page and section breaks
├── caption.go           # Auto-numbered captions
├── list.go              # Bullet and numbered lists
├── read.go              # Text extraction and search
├── replace.go           # Find and replace operations
├── properties.go        # Document properties
├── helpers.go           # Shared utility functions
├── utils.go             # ZIP and file utilities
├── types.go             # Shared type definitions
├── constants.go         # Constants and enums
├── errors.go            # Structured error types
├── doc.go               # Package-level documentation
├── *_test.go            # Unit and golden file tests
├── examples/            # Example programs
└── LICENSE              # MIT License
```

## Examples

Check the `/examples` directory for complete working examples:

- `example_all_features.go` - **Comprehensive demo** of every feature
- `example_toc_watermark.go` - TOC, watermarks, page numbers, styles, footnotes
- `example_chart_insert.go` - Creating charts from scratch
- `example_extended_chart.go` - Extended chart options
- `example_bookmarks.go` - Bookmark creation and internal navigation
- `example_table.go` - Table creation with styling
- `example_table_widths.go` - Table column width control
- `example_table_row_heights.go` - Table row height control
- `example_table_named_styles.go` - Named table styles
- `example_table_orientation.go` - Table with page orientation
- `example_paragraph.go` - Text and heading insertion
- `example_image.go` - Image insertion with proportional sizing
- `example_breaks.go` - Page and section breaks
- `example_captions.go` - Auto-numbered captions
- `example_lists.go` - Bullet and numbered lists
- `example_page_layout.go` - Page layout configuration
- `example_properties.go` - Document properties
- `example_blank_and_properties.go` - Blank document creation and full properties CRUD
- `example_multi_subsystem.go` - Combined operations
- `example_with_template.go` - Template-based generation
- `example_conditional_cell_colors.go` - Conditional cell formatting

Run any example:
```bash
go run examples/example_all_features.go template.docx output.docx
```

## Testing

```bash
# Run all tests
go test ./...

# Run with verbose output
go test -v ./...

# Run specific test
go test -run TestInsertTable ./...

# Generate coverage report
go test -cover ./...

# Run golden file tests only
go test -run TestGolden ./...
```

The test suite includes:
- Unit tests for all public and internal functions
- Golden file tests that verify XML output against expected strings
- Validation tests for error handling and edge cases

## Requirements

- Go 1.26 or later
- No external dependencies (uses only standard library)

## How It Works

DOCX files are ZIP archives containing XML files. This library:
1. Extracts the DOCX archive to a temporary directory
2. Parses and modifies XML files using Go's `encoding/xml` and string manipulation
3. Updates relationships (`_rels/*.rels`) and content types (`[Content_Types].xml`)
4. Manages embedded Excel workbooks for chart data
5. Re-packages everything into a new DOCX file

**Reliability details:**
- All in-place XML writes use an atomic write-then-rename strategy so a crash mid-write never leaves a corrupt file visible to readers
- The ZIP extractor enforces a 256 MiB per-file cap to guard against zip-bomb payloads
- XML escaping and unescaping use the stdlib `encoding/xml` codec throughout (no `html` package dependency)

### Concurrency Model

- `Updater` instances are isolated by temp directory, so using one `Updater` per goroutine/request is safe.
- A single `Updater` instance is **not** goroutine-safe; do not call methods on the same instance concurrently.
- Always call `defer u.Cleanup()` immediately after construction to avoid temp file leaks.

### OOXML Compatibility Notes

- `UpdateChart` requires `<c:externalData r:id="..."/>` and a valid `word/charts/_rels/chartN.xml.rels` target to an embedded workbook.
- Chart indices and delete/update indices are 1-based.
- Scatter chart updates expect numeric category values (used as X values).
- Inserted chart workbooks align series headers/data columns with chart formulas (`B..` for regular charts, `C..` when scatter X values occupy column `B`).
- Line chart series emit `c:marker` in schema-compliant `c:ser` child order for Microsoft 365 validation.
- `GetChartData` reads chart titles from both rich text (`a:t`) and value (`c:v`) title representations.
- Updating chart data rewrites chart series and worksheet `sheetData`; chart/workbook-level formatting in embedded workbooks is intentionally not preserved.
- Namespace prefix changes in output XML are expected and valid as long as namespace URIs remain correct.

## Roadmap

- [x] Chart updates and insertion (column, bar, line, pie, area, scatter)
- [x] Table insertion with cell merging
- [x] Paragraph and text insertion with formatting
- [x] Image insertion with proportional sizing
- [x] Embedded OLE objects (Excel workbooks, interactive double-click)
- [x] Header/footer manipulation
- [x] Table of Contents generation
- [x] Custom styles API
- [x] Watermarks
- [x] Page number control
- [x] Footnotes and endnotes
- [x] Comments
- [x] Track changes (insertions and deletions)
- [x] Delete operations
- [x] `io.Reader`/`io.Writer` support
- [x] Blank document creation (`NewBlank`, `NewFromBytes`)
- [x] Full properties CRUD (core, app, custom — get and set)
- [x] Expanded app properties (template, statistics, hyperlink base)
- [x] Golden file tests
- [ ] Content controls (structured document tags)
- [ ] Digital signatures
- [ ] Performance optimizations for large documents
- [ ] `context.Context` support for cancellation/timeouts

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Write tests for your changes
4. Commit your changes (`git commit -m 'Add amazing feature'`)
5. Push to the branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with Go's standard library
- Follows OpenXML specifications for DOCX manipulation
- Inspired by the need for programmatic Word document generation in Go

## Support

- 📫 Report issues on [GitHub Issues](https://github.com/falcomza/go-docx/issues)
- ⭐ Star this repo if you find it useful
- 🔧 Contributions and feedback are always welcome

---

Made with ❤️ by [falcomza](https://github.com/falcomza)

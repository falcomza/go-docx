# DOCX Updater

[![Go Version](https://img.shields.io/badge/Go-1.26+-00ADD8?style=flat&logo=go)](https://go.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful Go library for programmatically manipulating Microsoft Word (DOCX) documents. Update charts, insert tables, add paragraphs, generate captions, and more—all with a clean, idiomatic Go API.

## Features

🎯 **Comprehensive DOCX Manipulation**
- **Chart Updates**: Modify existing chart data with automatic Excel workbook synchronization
- **Chart Insertion**: Create professional charts from scratch (bar, line, scatter, and more)
- **Multi-Chart Workflows**: Insert multiple charts programmatically for bulk report generation
- **Table Creation**: Insert formatted tables with custom styles, borders, and row heights
- **Paragraph Insertion**: Add styled text with headings, alignment, list support, and robust anchor positioning
- **Image Insertion**: Add images with automatic proportional sizing and flexible positioning
- **Page & Section Breaks**: Control document flow with page and section breaks
- **Auto-Captions**: Generate auto-numbered captions using Word's SEQ fields for tables and charts
- **Text Find & Replace**: Search and replace text with regex support throughout documents
- **Read Operations**: Extract text from paragraphs, tables, headers, and footers
- **Hyperlinks**: Insert external URLs and internal document links
- **Bookmarks**: Create, manage, and reference bookmarks for internal navigation and TOC
- **Headers & Footers**: Professional document headers and footers with automatic page numbering
- **Document Properties**: Set core properties (Title, Author, Keywords), app properties (Company, Manager), and custom properties

🛠️ **Advanced Features**
- XML-based chart parsing using Go's `encoding/xml`
- Automatic Excel formula range adjustment
- Shared string table support for Excel workbooks
- Namespace-agnostic XML processing
- Full OpenXML relationship and content type management
- Strict workbook resolution via explicit relationships
- Structured error types for better error handling
- Case-sensitive and case-insensitive text operations
- Whole word matching and regex pattern support

## Installation

```bash
go get github.com/falcomza/go-docx
```

## Quick Start

```go
package main

import (
    "log"
    updater "github.com/falcomza/go-docx"
)

func main() {
    // Open existing DOCX
    u, err := updater.New("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer u.Cleanup()

    // Update a chart
    chartData := updater.ChartData{
        Categories: []string{"Q1", "Q2", "Q3", "Q4"},
        Series: []updater.SeriesData{
            {Name: "Revenue", Values: []float64{100, 150, 120, 180}},
            {Name: "Costs", Values: []float64{80, 90, 85, 95}},
        },
    }
    u.UpdateChart(1, chartData)

    // Add a table with caption
    u.InsertTable(updater.TableOptions{
        Columns: []updater.ColumnDefinition{
            {Title: "Product"},
            {Title: "Sales"},
            {Title: "Growth"},
        },
        Rows: [][]string{
            {"Product A", "$1.2M", "+15%"},
            {"Product B", "$980K", "+8%"},
        },
        TableStyle: updater.TableStyleGridAccent1,
        Position:   updater.PositionEnd,
        HeaderBold: true,
    })
    u.AddCaption(updater.CaptionOptions{
        Type:     updater.CaptionTypeTable,
        Label:    "Table",
        Position: updater.PositionEnd,
    })

    // Save result
    if err := u.Save("output.docx"); err != nil {
        log.Fatal(err)
    }
}
```

## Usage Examples

### Updating Chart Data

Update existing charts in a DOCX template:

```go
u, _ := updater.New("template.docx")
defer u.Cleanup()

data := updater.ChartData{
    Categories: []string{"Jan", "Feb", "Mar", "Apr"},
    Series: []updater.SeriesData{
        {Name: "Sales", Values: []float64{250, 300, 275, 350}},
    },
}

u.UpdateChart(1, data) // Update first chart (1-based index)
u.Save("updated.docx")
```

### Inserting New Charts

Create charts from scratch:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

chartOptions := updater.ChartOptions{
    Title:      "Quarterly Revenue",
    ChartKind:  updater.ChartKindColumn,
    Position:   updater.PositionEnd,
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []updater.SeriesOptions{
        {Name: "2025", Values: []float64{100, 120, 110, 130}},
        {Name: "2026", Values: []float64{110, 130, 125, 145}},
    },
}

u.InsertChart(chartOptions)
u.Save("with_chart.docx")
```

### Creating Tables

Insert styled tables with comprehensive formatting:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

u.InsertTable(updater.TableOptions{
    Columns: []updater.ColumnDefinition{
        {Title: "Product"},
        {Title: "Q1"},
        {Title: "Q2"},
        {Title: "Q3"},
        {Title: "Q4"},
    },
    Rows: [][]string{
        {"Product A", "$120K", "$135K", "$128K", "$150K"},
        {"Product B", "$98K", "$105K", "$112K", "$118K"},
        {"Product C", "$85K", "$92K", "$88K", "$95K"},
    },
    TableStyle: updater.TableStyleGridAccent1,
    Position:   updater.PositionEnd,
    HeaderBold: true,
    RowHeight:  280, // In twips (1/1440 inch)
})
u.Save("with_table.docx")
```

### Adding Paragraphs

Insert formatted text with various styles:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Add heading
u.AddHeading(1, "Executive Summary", updater.PositionEnd)

// Add normal text
u.AddText("This quarter showed strong growth across all regions.", updater.PositionEnd)

// Add formatted paragraph
u.InsertParagraph(updater.ParagraphOptions{
    Text:      "Important: Review required",
    Bold:      true,
    Italic:    true,
    Underline: true,
    Alignment: updater.ParagraphAlignCenter,
    Position:  updater.PositionEnd,
})

// Add paragraph after anchor text (works across split Word runs)
u.InsertParagraph(updater.ParagraphOptions{
    Text:     "Follow-up details",
    Position: updater.PositionAfterText,
    Anchor:   "Executive Summary",
})

// Newlines and tabs are emitted as <w:br/> and <w:tab/>
u.InsertParagraph(updater.ParagraphOptions{
    Text:     "Line 1\nLine 2\tTabbed",
    Position: updater.PositionEnd,
})

u.Save("with_paragraphs.docx")
```

**Paragraph Notes:**
- Supports alignment via `ParagraphAlignLeft`, `ParagraphAlignCenter`, `ParagraphAlignRight`, `ParagraphAlignJustify`
- `PositionEnd` insertion is section-safe (`w:sectPr` remains the final element in `<w:body>`)
- Anchor matching for `PositionAfterText` / `PositionBeforeText` is paragraph-aware and resilient to split runs
- Anchor matching also tolerates normalized whitespace differences (spaces/newlines/tabs)

### Inserting Images

Add images with automatic proportional sizing:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Insert image with width only - height calculated proportionally
u.InsertImage(updater.ImageOptions{
    Path:     "images/logo.png",
    Width:    400,  // pixels
    AltText:  "Company Logo",
    Position: updater.PositionEnd,
})

// Insert image with height only - width calculated proportionally
u.InsertImage(updater.ImageOptions{
    Path:     "images/chart.jpg",
    Height:   300,  // pixels
    AltText:  "Chart Illustration",
    Position: updater.PositionEnd,
})

// Insert image with both dimensions (may distort if not proportional)
u.InsertImage(updater.ImageOptions{
    Path:     "images/photo.png",
    Width:    500,
    Height:   400,
    Position: updater.PositionEnd,
})

// Insert image with actual file dimensions
u.InsertImage(updater.ImageOptions{
    Path:     "images/screenshot.png",
    AltText:  "Application Screenshot",
    Position: updater.PositionEnd,
})

// Insert image after specific text
u.InsertImage(updater.ImageOptions{
    Path:     "images/diagram.png",
    Width:    600,
    Position: updater.PositionAfterText,
    Anchor:   "See diagram below",
})

u.Save("with_images.docx")
```

**Proportional Sizing:**

- Specify only `Width`: Height calculated automatically
- Specify only `Height`: Width calculated automatically
- Specify both: Used as-is (may distort)
- Specify neither: Uses actual image dimensions

**Supported Formats:**

- PNG, JPEG, GIF, BMP, TIFF

**Image Captions:**

Images support auto-numbered captions using Word's SEQ fields:

```go
// Insert image with auto-numbered caption (Figure 1, Figure 2, etc.)
u.InsertImage(updater.ImageOptions{
    Path:     "images/chart.png",
    Width:    500,
    AltText:  "Sales Chart",
    Position: updater.PositionEnd,
    Caption: &updater.CaptionOptions{
        Type:        updater.CaptionFigure,
        Description: "Q1 Sales Performance",
        AutoNumber:  true,
        Position:    updater.CaptionAfter, // Caption below image (default)
    },
})

// Image with caption above
u.InsertImage(updater.ImageOptions{
    Path:     "images/diagram.png",
    Height:   350,
    Position: updater.PositionEnd,
    Caption: &updater.CaptionOptions{
        Type:        updater.CaptionFigure,
        Description: "Process Flow Diagram",
        AutoNumber:  true,
        Position:    updater.CaptionBefore, // Caption above image
        Alignment:   updater.CellAlignCenter, // Center the caption
    },
})
```

### Page and Section Breaks

Control document flow and layout with breaks:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Insert a page break to start new content on next page
u.InsertPageBreak(updater.BreakOptions{
    Position: updater.PositionEnd,
})

// Insert page break after specific text
u.InsertPageBreak(updater.BreakOptions{
    Position: updater.PositionAfterText,
    Anchor:   "End of Chapter 1",
})

// Insert section break (next page) - allows different page settings
u.InsertSectionBreak(updater.BreakOptions{
    Position:    updater.PositionEnd,
    SectionType: updater.SectionBreakNextPage,
    PageLayout:  updater.PageLayoutA3Landscape(),
})

// Insert continuous section break (same page, different formatting)
u.InsertSectionBreak(updater.BreakOptions{
    Position:    updater.PositionEnd,
    SectionType: updater.SectionBreakContinuous,
})

// Insert even/odd page section breaks (for double-sided printing)
u.InsertSectionBreak(updater.BreakOptions{
    Position:    updater.PositionEnd,
    SectionType: updater.SectionBreakEvenPage,
})

u.InsertSectionBreak(updater.BreakOptions{
    Position:    updater.PositionEnd,
    SectionType: updater.SectionBreakOddPage,
})

u.Save("with_breaks.docx")
```

**Section Break Types:**

- `SectionBreakNextPage` - Start new section on next page
- `SectionBreakContinuous` - Start new section on same page
- `SectionBreakEvenPage` - Start new section on next even page
- `SectionBreakOddPage` - Start new section on next odd page

**Use Cases:**

- Page breaks: Separate chapters, start appendices on new pages
- Section breaks: Different page orientations, margins, headers/footers per section
- Even/Odd breaks: Professional double-sided printing layouts

**Layout Helper Functions:**

- `PageLayoutLetterPortrait()` / `PageLayoutLetterLandscape()`
- `PageLayoutA4Portrait()` / `PageLayoutA4Landscape()`
- `PageLayoutA3Portrait()` / `PageLayoutA3Landscape()`
- `PageLayoutLegalPortrait()`

### Auto-Numbering Captions

Add captions with automatic sequential numbering:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Insert table
u.InsertTable(tableOptions)

// Add caption below the table
u.AddCaption(updater.CaptionOptions{
    Type:     updater.CaptionTypeTable,
    Label:    "Table",
    Text:     "Quarterly Sales Data",
    Position: updater.PositionEnd,
})

// Insert chart
u.InsertChart(chartOptions)

// Add caption below the chart
u.AddCaption(updater.CaptionOptions{
    Type:     updater.CaptionTypeChart,
    Label:    "Figure",
    Text:     "Revenue Trends 2025-2026",
    Position: updater.PositionEnd,
})

u.Save("with_captions.docx")
```

### Multiple Charts

Create multiple charts for bulk report generation:

```go
u, _ := updater.New("template.docx")
defer u.Cleanup()

// salesData is [][]updater.SeriesOptions, regions is [][]string
// Insert three charts with different data
for i := 0; i < 3; i++ {
    chartOptions := updater.ChartOptions{
        Position:   updater.PositionEnd,
        ChartKind:  updater.ChartKindColumn,
        Title:      fmt.Sprintf("Regional Report %d", i+1),
        Categories: regions[i],
        Series:     salesData[i], // salesData[i] is []updater.SeriesOptions
        ShowLegend: true,
    }
    u.InsertChart(chartOptions)
}

u.Save("multi_chart_report.docx")
```

### Text Find & Replace

Search and replace text throughout the document:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Simple case-insensitive replacement
opts := updater.DefaultReplaceOptions()
count, err := u.ReplaceText("{{name}}", "John Doe", opts)
fmt.Printf("Replaced %d occurrences\n", count)

// Case-sensitive whole word replacement
opts.MatchCase = true
opts.WholeWord = true
count, err = u.ReplaceText("API", "Application Programming Interface", opts)

// Replace with regex pattern
pattern := regexp.MustCompile(`\d{3}-\d{3}-\d{4}`) // Phone numbers
count, err = u.ReplaceTextRegex(pattern, "[REDACTED]", opts)

// Replace in specific locations
opts.InParagraphs = true
opts.InTables = true
opts.InHeaders = true  // Also replace in headers
opts.InFooters = true  // Also replace in footers

// Limit number of replacements
opts.MaxReplacements = 5  // Replace only first 5 occurrences

u.Save("replaced.docx")
```

### Read Operations

Extract and search for text in documents:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Get all text from document
text, err := u.GetText()
fmt.Println(text)

// Get text by paragraphs
paragraphs, err := u.GetParagraphText()
for i, para := range paragraphs {
    fmt.Printf("Paragraph %d: %s\n", i, para)
}

// Get text from tables
tables, err := u.GetTableText()
for i, table := range tables {
    fmt.Printf("Table %d:\n", i)
    for _, row := range table {
        fmt.Printf("  Row: %v\n", row)
    }
}

// Find all occurrences of text
opts := updater.DefaultFindOptions()
opts.MatchCase = false
matches, err := u.FindText("TODO:", opts)

for _, match := range matches {
    fmt.Printf("Found at paragraph %d: %s\n", match.Paragraph, match.Text)
    fmt.Printf("  Before: ...%s\n", match.Before)
    fmt.Printf("  After: %s...\n", match.After)
}

// Find with regex
opts.UseRegex = true
matches, err = u.FindText(`\b[A-Z]{2,}\b`, opts) // Find acronyms

// Limit search results
opts.MaxResults = 10  // Return only first 10 matches
```

### Hyperlinks

Insert clickable links to external URLs or internal bookmarks:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Insert external hyperlink
opts := updater.DefaultHyperlinkOptions()
opts.Position = updater.PositionEnd
opts.Tooltip = "Visit our website"

err := u.InsertHyperlink("Click here", "https://example.com", opts)

// Insert hyperlink after specific text
opts.Position = updater.PositionAfterText
opts.Anchor = "See our website"
err = u.InsertHyperlink("example.com", "https://example.com", opts)

// Customize hyperlink appearance
opts.Color = "FF0000"     // Red color
opts.Underline = true     // Underline (default)
opts.ScreenTip = "Link"   // Accessibility text

// Insert email link
err = u.InsertHyperlink("Contact Us", "mailto:info@example.com", opts)

// Insert internal link to bookmark
err = u.InsertInternalLink("Go to Summary", "summary_bookmark", opts)

u.Save("with_links.docx")
```

### Bookmarks

Create bookmarks to mark locations in your document and enable internal navigation:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Create an empty bookmark (position marker)
opts := updater.DefaultBookmarkOptions()
opts.Position = updater.PositionEnd
err := u.CreateBookmark("section_marker", opts)

// Create bookmark with text content
opts.Style = updater.StyleHeading1
err = u.CreateBookmarkWithText("executive_summary", "Executive Summary", opts)

// Wrap existing text in a bookmark
err = u.WrapTextInBookmark("key_finding", "important result")

// Create internal links to bookmarks
linkOpts := updater.DefaultHyperlinkOptions()
linkOpts.Position = updater.PositionBeginning
err = u.InsertInternalLink("Jump to Summary", "executive_summary", linkOpts)

// Position-based bookmark insertion
opts.Position = updater.PositionAfterText
opts.Anchor = "Chapter 3"
err = u.CreateBookmark("chapter3_bookmark", opts)

u.Save("with_bookmarks.docx")
```

**Bookmark Name Rules:**
- Must start with a letter
- Can contain letters, digits, and underscores
- No spaces or special characters (except underscore)
- Maximum 40 characters
- Cannot start with reserved prefixes (`_Toc`, `_Hlt`, `_Ref`, `_GoBack`)

**Common Use Cases:**
- Table of contents with clickable links
- Cross-references within documents
- Navigation between sections
- Marking important locations for reference

### Headers and Footers

Add professional headers and footers with automatic page numbering:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Create header with three-column layout
headerContent := updater.HeaderFooterContent{
    LeftText:   "Company Name",
    CenterText: "Confidential Report",
    RightText:  "Date: Feb 2026",
    PageNumber: false,
}

headerOpts := updater.DefaultHeaderOptions()
headerOpts.Type = updater.HeaderDefault
err := u.SetHeader(headerContent, headerOpts)

// Create footer with page numbers
footerContent := updater.HeaderFooterContent{
    CenterText:       "Page ",
    PageNumber:       true,
    PageNumberFormat: "X of Y",  // Shows "Page 1 of 10"
}

footerOpts := updater.DefaultFooterOptions()
err = u.SetFooter(footerContent, footerOpts)

// Different header for first page
headerOpts.Type = updater.HeaderFirst
headerOpts.DifferentFirst = true
firstPageHeader := updater.HeaderFooterContent{
    CenterText: "Title Page - No Header",
}
err = u.SetHeader(firstPageHeader, headerOpts)

// Different headers for odd/even pages (for double-sided printing)
headerOpts.DifferentOddEven = true

// Odd pages (right side)
headerOpts.Type = updater.HeaderDefault
oddHeader := updater.HeaderFooterContent{
    RightText: "Chapter 1",
}
err = u.SetHeader(oddHeader, headerOpts)

// Even pages (left side)
headerOpts.Type = updater.HeaderEven
evenHeader := updater.HeaderFooterContent{
    LeftText: "Chapter 1",
}
err = u.SetHeader(evenHeader, headerOpts)

// Add date field to footer
footerContent.Date = true
footerContent.DateFormat = "MMMM d, yyyy"  // "January 1, 2026"

u.Save("with_headers_footers.docx")
```

### Setting Document Properties

Set core, application, and custom document properties:

```go
u, _ := updater.New("template.docx")
defer u.Cleanup()

// Set core properties (visible in File > Info)
coreProps := updater.CoreProperties{
    Title:          "Q4 2026 Financial Report",
    Subject:        "Quarterly Financial Analysis",
    Creator:        "Finance Department",
    Keywords:       "finance, Q4, 2026, revenue, analysis",
    Description:    "Comprehensive financial report for Q4 2026",
    Category:       "Financial Reports",
    LastModifiedBy: "John Doe",
    Revision:       "2",
    Created:        time.Date(2026, 1, 1, 9, 0, 0, 0, time.UTC),
    Modified:       time.Now(),
}
err := u.SetCoreProperties(coreProps)

// Set application properties
appProps := updater.AppProperties{
    Company:     "TechVenture Inc",
    Manager:     "Sarah Williams",
    Application: "Microsoft Word",
    AppVersion:  "16.0000",
}
err = u.SetAppProperties(appProps)

// Set custom properties (for workflow automation, metadata tracking, etc.)
customProps := []updater.CustomProperty{
    {Name: "Department", Value: "Finance", Type: "lpwstr"},
    {Name: "FiscalYear", Value: 2026, Type: "i4"},
    {Name: "Quarter", Value: "Q4"},
    {Name: "Revenue", Value: 15750000.50, Type: "r8"},
    {Name: "IsApproved", Value: true, Type: "bool"},
    {Name: "ProjectCode", Value: "FIN-Q4-2026"},
    {Name: "ConfidentialityLevel", Value: "High"},
}
err = u.SetCustomProperties(customProps)

// Read core properties
props, err := u.GetCoreProperties()
fmt.Printf("Title: %s, Author: %s\n", props.Title, props.Creator)

u.Save("with_properties.docx")
```

**Property Types:**
- Core Properties: Title, Subject, Creator, Keywords, Description, Category, LastModifiedBy, Revision, Created, Modified
- App Properties: Company, Manager, Application, AppVersion
- Custom Properties: Any key-value pairs with types: string (`lpwstr`), integer (`i4`), float (`r8`), boolean (`bool`), date (`date`)

## API Overview

### Chart Operations
- `UpdateChart(index int, data ChartData)` - Update existing chart data
- `InsertChart(options ChartOptions)` - Create new chart from scratch

### Table Operations
- `InsertTable(options TableOptions)` - Insert formatted table with custom styling (columns and rows are fields inside `TableOptions`)

### Paragraph Operations
- `InsertParagraph(options ParagraphOptions)` - Insert styled paragraph
- `InsertParagraphs(paragraphs []ParagraphOptions)` - Insert multiple paragraphs
- `AddHeading(level int, text string, position InsertPosition)` - Insert heading paragraph
- `AddText(text string, position InsertPosition)` - Insert normal paragraph text
- `AddBulletItem(text string, level int, position InsertPosition)` - Insert bullet list item
- `AddNumberedItem(text string, level int, position InsertPosition)` - Insert numbered list item

### Image Operations
- `InsertImage(options ImageOptions)` - Insert image with proportional sizing

### Text Operations
- `ReplaceText(old, new string, options ReplaceOptions)` - Replace all text occurrences
- `ReplaceTextRegex(pattern *regexp.Regexp, replacement string, options ReplaceOptions)` - Replace using regex
- `GetText()` - Extract all text from document
- `GetParagraphText()` - Extract text from all paragraphs
- `GetTableText()` - Extract text from all tables
- `FindText(pattern string, options FindOptions)` - Find all occurrences with context

### Hyperlink Operations
- `InsertHyperlink(text, url string, options HyperlinkOptions)` - Insert external hyperlink
- `InsertInternalLink(text, bookmarkName string, options HyperlinkOptions)` - Insert internal link

### Bookmark Operations
- `CreateBookmark(name string, options BookmarkOptions)` - Create empty bookmark marker
- `CreateBookmarkWithText(name, text string, options BookmarkOptions)` - Create bookmark with text content
- `WrapTextInBookmark(name, anchorText string)` - Wrap existing text in bookmark

### Header & Footer Operations
- `SetHeader(content HeaderFooterContent, options HeaderOptions)` - Create/update header
- `SetFooter(content HeaderFooterContent, options FooterOptions)` - Create/update footer

### Properties Operations
- `SetCoreProperties(props CoreProperties)` - Set core document properties (Title, Author, etc.)
- `GetCoreProperties()` - Retrieve core document properties
- `SetAppProperties(props AppProperties)` - Set application properties (Company, Manager, etc.)
- `SetCustomProperties(properties []CustomProperty)` - Set custom properties with various types

### Break Operations
- `InsertPageBreak(options BreakOptions)` - Insert page break
- `InsertSectionBreak(options BreakOptions)` - Insert section break

### Caption Operations
- `AddCaption(options CaptionOptions)` - Insert auto-numbered caption

### Core Operations
- `New(filepath string) (*Updater, error)` - Open DOCX file
- `Save(outputPath string) error` - Save modified document
- `Cleanup()` - Clean up temporary files

## Project Structure

```
.
├── *.go                   # Core library (root level)
│   ├── chart_updater.go   # Main API
│   ├── chart.go           # Chart insertion
│   ├── chart_xml.go       # XML manipulation
│   ├── excel_handler.go   # Workbook updates
│   ├── table.go           # Table insertion
│   ├── paragraph.go       # Text insertion
│   ├── image.go           # Image insertion
│   ├── bookmark.go        # Bookmark management
│   ├── breaks.go          # Page and section breaks
│   ├── caption.go         # Caption generation
│   └── ...
├── *_test.go              # Unit tests (root level)
├── examples/              # Example programs
├── templates/             # Sample templates
└── LICENSE                # MIT License
```

## Examples

Check the `/examples` directory for complete working examples:

- `example_bookmarks.go` - Bookmark creation and internal navigation
- `example_chart_insert.go` - Creating charts from scratch
- `example_table.go` - Table creation with styling
- `example_paragraph.go` - Text and heading insertion
- `example_image.go` - Image insertion with proportional sizing
- `example_breaks.go` - Page and section breaks
- `example_captions.go` - Auto-numbered captions
- `example_multi_subsystem.go` - Combined operations
- `example_with_template.go` - Template-based generation

Run any example:
```bash
go run examples/example_table.go
```

## Testing

Run the comprehensive test suite:

```bash
# Run all tests
go test ./...

# Run specific test
go test -run TestInsertTable ./...

# Run with verbose output
go test -v ./...

# Generate coverage report
go test -cover ./...
```

## Requirements

- Go 1.26 or later
- No external dependencies (uses only standard library)

## How It Works

DOCX files are ZIP archives containing XML files. This library:
1. Extracts the DOCX archive to a temporary directory
2. Parses and modifies XML files using Go's `encoding/xml`
3. Updates relationships (`_rels/*.rels`) and content types
4. Manages embedded Excel workbooks for chart data
5. Re-packages everything into a new DOCX file

## Limitations

- Supports bar, column, line, pie, area, and scatter chart types
- Table styles are limited to predefined Word styles
- Performance depends on document size and complexity

## Roadmap

- [x] Image insertion support with proportional sizing
- [x] Header/footer manipulation
- [ ] Style customization API
- [ ] Performance optimizations for large documents

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

# DOCX Updater

[![Go Version](https://img.shields.io/badge/Go-1.23+-00ADD8?style=flat&logo=go)](https://go.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful Go library for programmatically manipulating Microsoft Word (DOCX) documents. Update charts, insert tables, add paragraphs, generate captions, and more—all with a clean, idiomatic Go API.

## Features

🎯 **Comprehensive DOCX Manipulation**
- **Chart Updates**: Modify existing chart data with automatic Excel workbook synchronization
- **Chart Insertion**: Create professional charts from scratch (bar, line, scatter, and more)
- **Chart Copying**: Duplicate existing charts programmatically for bulk report generation
- **Table Creation**: Insert formatted tables with custom styles, borders, and row heights
- **Paragraph Insertion**: Add styled text with headings, bold, italic, and underline formatting
- **Auto-Captions**: Generate auto-numbered captions using Word's SEQ fields for tables and charts

🛠️ **Advanced Features**
- XML-based chart parsing using Go's `encoding/xml`
- Automatic Excel formula range adjustment
- Shared string table support for Excel workbooks
- Namespace-agnostic XML processing
- Full OpenXML relationship and content type management
- Strict workbook resolution via explicit relationships

## Installation

```bash
go get github.com/falcomza/docx-update
```

## Quick Start

```go
package main

import (
    "log"
    updater "github.com/falcomza/docx-update/src"
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
    table := updater.TableData{
        Headers: []string{"Product", "Sales", "Growth"},
        Rows: [][]string{
            {"Product A", "$1.2M", "+15%"},
            {"Product B", "$980K", "+8%"},
        },
    }
    u.InsertTable(table, updater.TableOptions{
        Style:    updater.TableStyleGridTable4Accent1,
        Position: updater.PositionEnd,
    })
    u.AddCaption(updater.CaptionOptions{
        Type:       updater.CaptionTypeTable,
        Label:      "Table",
        Position:   updater.PositionEnd,
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

chartOptions := updater.ChartInsertOptions{
    Title:      "Quarterly Revenue",
    ChartType:  updater.ChartTypeColumn,
    Position:   updater.PositionEnd,
    Width:      6.0,  // inches
    Height:     4.0,  // inches
    Data: updater.ChartData{
        Categories: []string{"Q1", "Q2", "Q3", "Q4"},
        Series: []updater.SeriesData{
            {Name: "2025", Values: []float64{100, 120, 110, 130}},
            {Name: "2026", Values: []float64{110, 130, 125, 145}},
        },
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

table := updater.TableData{
    Headers: []string{"Product", "Q1", "Q2", "Q3", "Q4"},
    Rows: [][]string{
        {"Product A", "$120K", "$135K", "$128K", "$150K"},
        {"Product B", "$98K", "$105K", "$112K", "$118K"},
        {"Product C", "$85K", "$92K", "$88K", "$95K"},
    },
}

options := updater.TableOptions{
    Style:          updater.TableStyleGridTable4Accent1,
    Position:       updater.PositionEnd,
    HeaderBold:     true,
    Border:         true,
    RowHeights:     []int{300, 280, 280, 280}, // In twips (1/1440 inch)
}

u.InsertTable(table, options)
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
    Position:  updater.PositionEnd,
})

u.Save("with_paragraphs.docx")
```

### Auto-Numbering Captions

Add captions with automatic sequential numbering:

```go
u, _ := updater.New("document.docx")
defer u.Cleanup()

// Insert table
u.InsertTable(tableData, tableOptions)

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

### Copying Charts

Duplicate existing charts for bulk report generation:

```go
u, _ := updater.New("template.docx")
defer u.Cleanup()

// Copy chart 1 three times with different data
for i := 0; i < 3; i++ {
    data := updater.ChartData{
        Categories: regions[i],
        Series:     salesData[i],
    }
    u.CopyChart(1, data, updater.PositionEnd)
}

u.Save("multi_chart_report.docx")
```

## API Overview

### Chart Operations
- `UpdateChart(index int, data ChartData)` - Update existing chart data
- `InsertChart(options ChartInsertOptions)` - Create new chart from scratch
- `CopyChart(index int, data ChartData, position Position)` - Duplicate existing chart

### Table Operations
- `InsertTable(data TableData, options TableOptions)` - Insert formatted table
- Supports: custom styles, borders, row heights, column widths, alignments

### Paragraph Operations
- `AddText(text string, position Position)` - Insert plain text
- `AddHeading(level int, text string, position Position)` - Insert heading (1-6)
- `InsertParagraph(options ParagraphOptions)` - Insert formatted paragraph
- Supports: bold, italic, underline, custom styles

### Caption Operations
- `AddCaption(options CaptionOptions)` - Add auto-numbered caption
- Uses Word's SEQ fields for automatic numbering
- Supports both tables and charts

### Core Operations
- `New(filepath string) (*Updater, error)` - Open DOCX file
- `Save(outputPath string) error` - Save modified document
- `Cleanup()` - Clean up temporary files

## Project Structure

```
.
├── src/                    # Core library
│   ├── chart_updater.go   # Main API
│   ├── chart.go           # Chart insertion
│   ├── chart_copy.go      # Chart duplication
│   ├── chart_xml.go       # XML manipulation
│   ├── excel_handler.go   # Workbook updates
│   ├── table.go           # Table insertion
│   ├── paragraph.go       # Text insertion
│   ├── caption.go         # Caption generation
│   └── ...
├── tests/                 # Unit tests
├── examples/              # Example programs
└── templates/             # Sample templates
```

## Examples

Check the `/examples` directory for complete working examples:

- `example_chart_insert.go` - Creating charts from scratch
- `example_table.go` - Table creation with styling
- `example_paragraph.go` - Text and heading insertion
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
go test ./tests/...

# Run specific test
go test ./tests/ -run TestInsertTable

# Run with verbose output
go test -v ./tests/...

# Generate coverage report
go test -cover ./tests/...
```

## Requirements

- Go 1.23 or later
- No external dependencies (uses only standard library)

## How It Works

DOCX files are ZIP archives containing XML files. This library:
1. Extracts the DOCX archive to a temporary directory
2. Parses and modifies XML files using Go's `encoding/xml`
3. Updates relationships (`_rels/*.rels`) and content types
4. Manages embedded Excel workbooks for chart data
5. Re-packages everything into a new DOCX file

## Limitations

- Currently supports bar, line, and scatter chart types
- Table styles are limited to predefined Word styles
- Images in charts are not yet supported
- Performance depends on document size and complexity

## Roadmap

- [ ] Add more chart types (pie, area, combo charts)
- [ ] Image insertion support
- [ ] Header/footer manipulation
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

This project is licensed under the MIT License - see below for details:

```
MIT License

Copyright (c) 2026 falcomza

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Acknowledgments

- Built with Go's standard library
- Follows OpenXML specifications for DOCX manipulation
- Inspired by the need for programmatic Word document generation in Go

## Support

- 📫 Report issues on [GitHub Issues](https://github.com/falcomza/docx-update/issues)
- ⭐ Star this repo if you find it useful
- 🔧 Contributions and feedback are always welcome

---

Made with ❤️ by [falcomza](https://github.com/falcomza)

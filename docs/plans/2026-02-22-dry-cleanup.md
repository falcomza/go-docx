# DRY Cleanup & README Fix Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Remove dead code, merge the duplicate chart API into one, fix all README inaccuracies, and align the Go version to 1.26 everywhere.

**Architecture:** Single `InsertChart(ChartOptions)` replaces both `InsertChart` (simple) and `InsertChartExtended` (full). `ChartOptions.Series` becomes `[]SeriesOptions` (superset of old `[]SeriesData`). Optional extended fields (`CategoryAxis`, `ValueAxis`, `Legend`, `DataLabels`, `Properties`, `BarChartOptions`) are added to `ChartOptions` as nil-able pointers. The extended XML generation path becomes the only path. Dead functions and vars are deleted.

**Tech Stack:** Go 1.26, `encoding/xml`, `archive/zip`, standard library only.

---

### Task 1: Go version alignment

**Files:**
- Modify: `go.mod`
- Modify: `README.md` (Requirements section)
- Modify: `CLAUDE.md` (Go version line)

**Step 1: Update go.mod**

Change `go 1.25` → `go 1.26`.

**Step 2: Update README.md Requirements section**

Change `Go 1.23 or later` → `Go 1.26 or later`.

**Step 3: Verify**

```bash
go build ./...
```
Expected: clean build, no errors.

**Step 4: Commit**

```bash
git add go.mod README.md CLAUDE.md
git commit -m "chore: align Go version to 1.26 across go.mod, README, CLAUDE.md"
```

---

### Task 2: Remove dead code — constants

**Files:**
- Modify: `constants.go`

**Context:** `chartRelPatternTemplate` (line 55) and `workbookNumberPattern` (line 58) are declared but never referenced outside this file.

**Step 1: Delete both vars and their comments from `constants.go`**

Remove lines 53–58:
```go
// chartRelPatternTemplate is a format string for matching specific chart relationships
// Use with fmt.Sprintf to insert the chart index
chartRelPatternTemplate = `Id="(rId[0-9]+)"[^>]*Target="charts/chart%d\.xml"`

// workbookNumberPattern matches numeric suffixes in workbook filenames
workbookNumberPattern = regexp.MustCompile(`^(.+?)(\d+)$`)
```

If the `var (...)` block becomes empty as a result, remove the block. Otherwise keep the remaining vars.

**Step 2: Remove `"regexp"` import from `constants.go` if `workbookNumberPattern` was the only regexp usage**

Check: `chartFilePattern`, `imageFilePattern`, `docPrIDPattern`, `bookmarkIDPattern`, `relIDPattern`, `textRunPattern`, `textContentPattern` all use `regexp.MustCompile`. Keep the import.

**Step 3: Verify**

```bash
go build ./...
go test ./...
```
Expected: clean build and all tests pass.

**Step 4: Commit**

```bash
git add constants.go
git commit -m "chore: remove unused chartRelPatternTemplate and workbookNumberPattern vars"
```

---

### Task 3: Remove dead code — copyFile

**Files:**
- Modify: `utils.go`

**Context:** `copyFile` (lines 134–157) is declared but never called.

**Step 1: Delete the `copyFile` function and its doc comment**

Remove from `// copyFile copies a file...` through the closing `}`.

**Step 2: Verify**

```bash
go build ./...
go test ./...
```
Expected: clean.

**Step 3: Commit**

```bash
git add utils.go
git commit -m "chore: remove unused copyFile function"
```

---

### Task 4: Remove dead code — generateChartDrawingXML

**Files:**
- Modify: `chart.go`

**Context:** `generateChartDrawingXML` (lines 1475–1488) is declared as a method on `*Updater` but never called. The used path goes through `generateChartDrawingWithSize`.

**Step 1: Delete the method and its doc comment**

Remove `generateChartDrawingXML` from `chart.go`.

**Step 2: Verify**

```bash
go build ./...
go test ./...
```
Expected: clean.

**Step 3: Commit**

```bash
git add chart.go
git commit -m "chore: remove unused generateChartDrawingXML method"
```

---

### Task 5: Fix unused parameter in insertCaptionWithElement

**Files:**
- Modify: `caption.go`
- Modify: `chart.go` (two call sites)
- Modify: `image.go` (one call site)
- Modify: `table.go` (one call site)

**Context:** `insertCaptionWithElement(docXML, captionXML, elementXML []byte, position CaptionPosition)` — the `docXML` first parameter is never used in the function body. The callers pass `raw` (the doc bytes) as the first argument unnecessarily.

**Step 1: Remove `docXML` from the signature in `caption.go:167`**

```go
// Before
func insertCaptionWithElement(docXML, captionXML, elementXML []byte, position CaptionPosition) []byte {

// After
func insertCaptionWithElement(captionXML, elementXML []byte, position CaptionPosition) []byte {
```

**Step 2: Update all four call sites**

`chart.go:600`:
```go
// Before
contentToInsert = insertCaptionWithElement(raw, captionXML, drawingXML, opts.Caption.Position)
// After
contentToInsert = insertCaptionWithElement(captionXML, drawingXML, opts.Caption.Position)
```

`chart.go:1408`:
```go
// Before
contentToInsert = insertCaptionWithElement(raw, captionXML, drawing, opts.Caption.Position)
// After
contentToInsert = insertCaptionWithElement(captionXML, drawing, opts.Caption.Position)
```

`image.go:97`:
```go
// Before
contentToInsert = insertCaptionWithElement(raw, captionXML, imageXML, opts.Caption.Position)
// After
contentToInsert = insertCaptionWithElement(captionXML, imageXML, opts.Caption.Position)
```

`table.go:701`:
```go
// Before
contentToInsert = insertCaptionWithElement(docXML, captionXML, tableXML, opts.Caption.Position)
// After
contentToInsert = insertCaptionWithElement(captionXML, tableXML, opts.Caption.Position)
```

**Step 3: Verify**

```bash
go build ./...
go test ./...
```
Expected: clean.

**Step 4: Commit**

```bash
git add caption.go chart.go image.go table.go
git commit -m "fix: remove unused docXML parameter from insertCaptionWithElement"
```

---

### Task 6: Merge dual chart API — extend ChartOptions

**Files:**
- Modify: `chart.go`

**Context:** `ChartOptions` is the public struct for `InsertChart`. It needs the optional fields that `ExtendedChartOptions` has. `ExtendedChartOptions` will be deleted in Task 7.

**Step 1: Add optional extended fields to `ChartOptions`**

After the existing `Caption *CaptionOptions` field, add:
```go
// Extended axis customization (nil = auto defaults)
CategoryAxis *AxisOptions
ValueAxis    *AxisOptions

// Legend customization (nil = show legend on right)
Legend *LegendOptions

// Default data labels for all series (nil = no labels)
DataLabels *DataLabelOptions

// Chart-level rendering properties (nil = library defaults)
Properties *ChartProperties

// Bar/column-specific options (nil = clustered column defaults)
BarChartOptions *BarChartOptions

// If true, title overlays the chart area
TitleOverlay bool
```

**Step 2: Change `ChartOptions.Series` type from `[]SeriesData` to `[]SeriesOptions`**

```go
// Before
Series []SeriesData

// After
Series []SeriesOptions // Use SeriesOptions for per-series color, smooth, marker control
```

`SeriesOptions` is a superset of `SeriesData` (Name, Values, Color + InvertIfNegative, Smooth, ShowMarkers, DataLabels). All downstream functions only access `.Name` and `.Values` (for the workbook), so this is backward compatible at the data level.

**Step 3: Verify the change compiles (tests will fail until Task 8 updates them)**

```bash
go build ./...
```

---

### Task 7: Merge dual chart API — unify functions

**Files:**
- Modify: `chart.go`

This is the largest task. Work through it section by section.

**Step 1: Merge validate functions**

Delete `validateChartOptions` (lines ~107–126). It's a strict subset of `validateExtendedChartOptions`.

Rename `validateExtendedChartOptions` → `validateChartOptions` and change its parameter type from `ExtendedChartOptions` to `ChartOptions`.

Update the field names accessed (they're identical: `Categories`, `Series`, `CategoryAxis`, `ValueAxis`, `BarChartOptions`).

**Step 2: Merge defaults functions**

Delete `applyChartDefaults` (lines ~128–143). It's a subset of `applyExtendedChartDefaults`.

Rename `applyExtendedChartDefaults` → `applyChartDefaults` and change its parameter/return type from `ExtendedChartOptions` to `ChartOptions`.

All field accesses are identical after the type change.

**Step 3: Rename generateExtendedChartXML → generateChartXML**

Delete old `generateChartXML` (lines ~161–...). Rename `generateExtendedChartXML` → `generateChartXML`, changing parameter type to `ChartOptions`.

Rename internal helpers:
- `generateExtendedBarChartXML` → `generateBarChartXML`
- `generateExtendedLineChartXML` → `generateLineChartXML`
- `generateExtendedPieChartXML` → `generatePieChartXML`
- `generateExtendedAreaChartXML` → `generateAreaChartXML`

All parameter types change from `ExtendedChartOptions` to `ChartOptions`.

**Step 4: Update createChartXML**

`createChartXML` calls `generateChartXML`. Since that now refers to the (renamed) extended function, no logic change needed — just ensure the parameter type is `ChartOptions`.

Delete `createExtendedChartXML` (now redundant with `createChartXML`).

**Step 5: Merge drawing insertion**

Delete `insertChartDrawing` (the simple version). Rename `insertExtendedChartDrawing` → `insertChartDrawing`, change parameter type to `ChartOptions`.

**Step 6: Update createEmbeddedWorkbook**

`generateSheetXML` accesses `series.Name` and `series.Values` — both present on `SeriesOptions`. No logic change needed, the type is already `ChartOptions` whose `Series` is now `[]SeriesOptions`.

**Step 7: Delete InsertChartExtended and helpers**

Delete these functions:
- `InsertChartExtended`
- `convertToChartOptions`

**Step 8: Rewrite InsertChart to use unified path**

`InsertChart` now calls:
```
validateChartOptions  → merged validator
applyChartDefaults    → merged defaults (returns ChartOptions)
createChartXML        → uses generateChartXML (extended path)
createEmbeddedWorkbook
createChartRelationships
addChartRelationship
insertChartDrawing    → unified drawing insertion
addContentTypeOverride
```

**Step 9: Verify**

```bash
go build ./...
```

Expected: compile errors only in test files (fixed in Task 8).

---

### Task 8: Update tests after chart API merge

**Files:**
- Modify: `chart_extended_test.go`
- Modify: `caption_test.go`
- Modify: `comprehensive_demo_test.go`
- Modify: `chart_insert_test.go`
- Modify: `new_features_test.go`
- Modify: any other test file using `SeriesData` in `ChartOptions` context

**Step 1: Update chart_extended_test.go**

Replace all `ExtendedChartOptions{...}` → `ChartOptions{...}`.
Replace all `[]SeriesOptions{...}` references in opts — these stay `[]SeriesOptions` since `ChartOptions.Series` is now `[]SeriesOptions`.
Replace `InsertChartExtended` calls → `InsertChart`.
Replace `validateExtendedChartOptions` calls → `validateChartOptions`.
Replace `applyExtendedChartDefaults` calls → `applyChartDefaults`.
Replace `generateExtendedChartXML` calls → `generateChartXML`.

**Step 2: Update caption_test.go and other tests**

Find all `Series: []godocx.SeriesData{...}` inside `ChartOptions` and change to `[]godocx.SeriesOptions{...}`.

The struct literal fields `Name` and `Values` are the same on both types — only the type name changes.

**Step 3: Verify**

```bash
go test ./... -v 2>&1 | tail -20
```
Expected: all tests pass.

**Step 4: Commit**

```bash
git add chart.go chart_extended_test.go caption_test.go comprehensive_demo_test.go chart_insert_test.go new_features_test.go
git commit -m "refactor: merge InsertChartExtended into InsertChart, unify ChartOptions"
```

---

### Task 9: Delete ExtendedChartOptions type

**Files:**
- Modify: `chart_extended.go`

**Context:** `chart_extended.go` contains `ExtendedChartOptions` struct (lines 162–201) and all the supporting types (`ChartStyle`, `DataLabelOptions`, `AxisOptions`, `LegendOptions`, `SeriesOptions`, `ChartProperties`, `BarChartOptions`, etc.) that are still used.

**Step 1: Delete only the `ExtendedChartOptions` struct definition** (lines 162–201 in the original, the block starting with `// ExtendedChartOptions defines...`).

Keep all other types: `ChartStyle`, `DataLabelPosition`, `AxisPosition`, `TickMark`, `TickLabelPosition`, `BarGrouping`, `BarDirection`, `DataLabelOptions`, `AxisOptions`, `LegendOptions`, `SeriesOptions`, `ChartProperties`, `BarChartOptions`.

**Step 2: Verify**

```bash
go build ./...
go test ./...
```
Expected: all clean.

**Step 3: Commit**

```bash
git add chart_extended.go
git commit -m "chore: delete ExtendedChartOptions type, absorbed into ChartOptions"
```

---

### Task 10: Fix README.md

**Files:**
- Modify: `README.md`

Fix every inaccuracy found in audit:

**1. Go badge and Requirements**
- Badge: `Go-1.23+` → `Go-1.26+`
- Requirements section: `Go 1.23 or later` → `Go 1.26 or later`

**2. Quick Start — table**
The current Quick Start shows a fictional `TableData{Headers: ..., Rows: ...}` and `InsertTable(table, options)` two-parameter call. Replace with the real API:
```go
// Old (fictional)
table := updater.TableData{
    Headers: []string{"Product", "Sales", "Growth"},
    Rows: [][]string{...},
}
u.InsertTable(table, updater.TableOptions{...})

// Real API
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
```

**3. Quick Start — chart insertion**
`ChartInsertOptions` doesn't exist → `ChartOptions`. Also `Series []SeriesData` → `Series []SeriesOptions` in InsertChart context.

**4. Creating Tables example**
`TableStyleGridTable4Accent1` → `TableStyleGridAccent1`.
`InsertTable(table, options)` two-arg call → `InsertTable(options)` single-arg.

**5. Inserting New Charts example**
`ChartInsertOptions` → `ChartOptions`. Change `Series []SeriesData{...}` to `Series []SeriesOptions{...}`.

**6. Multiple Charts example**
`ChartOptions.Series` field type: change `[]updater.SeriesData` to `[]updater.SeriesOptions` in example.

**7. API Overview — Chart Operations**
Remove `InsertChartExtended` line. Update `InsertChart` signature note.
Add `InsertChart(options ChartOptions)` note: "Supports full axis, legend, data label, and series customization via optional fields."

**8. API Overview — Table Operations**
`InsertTable(options TableOptions)` — note that columns and rows are inside `TableOptions`.

**9. Limitations section**
Change:
```
- Currently supports bar, line, and scatter chart types
```
To:
```
- Supports bar, column, line, pie, area, and scatter chart types
```

**10. Roadmap section**
Remove `- [ ] Add more chart types (pie, area, combo charts)` — done.
Mark `- [ ] Header/footer manipulation` as `- [x] Header/footer manipulation` — done.

**11. Testing section**
```bash
# Before (wrong path)
go test ./tests/...
go test ./tests/ -run TestInsertTable
go test -v ./tests/...
go test -cover ./tests/...

# After (correct)
go test ./...
go test -run TestInsertTable ./...
go test -v ./...
go test -cover ./...
```

**12. Updater doc comment and struct description**
The `Updater` struct comment in `chart_updater.go` says `"updates chart caches and embedded workbook data"`. Change to `"manages a DOCX document for reading and writing"`.
`New()` doc comment similarly: change from chart-specific description to general purpose.

**Step: Verify the examples in README actually match**

For each code block in README, check:
- Type names exist in codebase
- Function signatures match
- Constants exist

```bash
grep -n "TableStyleGridTable4Accent1\|ChartInsertOptions\|TableData{" README.md
```
Expected: no output (all fictional names removed).

**Commit**

```bash
git add README.md chart_updater.go
git commit -m "docs: fix README accuracy — correct types, signatures, constants, test paths, roadmap"
```

---

### Task 11: Final verification

**Step 1: Full test suite**

```bash
go test ./... -v 2>&1 | grep -E "^(=== RUN|--- (PASS|FAIL|SKIP)|ok|FAIL)" | tail -40
```
Expected: all PASS, no FAIL.

**Step 2: No remaining fictional type references in README**

```bash
grep -n "TableData\|ChartInsertOptions\|TableStyleGridTable4Accent1\|ExtendedChartOptions\|InsertChartExtended\|SeriesData" README.md
```
Expected: no output.

**Step 3: No dead code warnings**

```bash
go vet ./...
```
Expected: clean.

**Step 4: Build examples**

```bash
for f in examples/*.go; do go build -o /dev/null "$f" 2>&1 && echo "OK: $f" || echo "FAIL: $f"; done
```
Expected: all OK (or pre-existing failures in examples that used the fictional API).

**Step 5: Push**

```bash
git push
```

//go:build ignore

// changelog_updater processes a folder of DOCX files:
//  1. Sets each document's core-properties Subject to "<original subject> YYYY"
//     where YYYY is the current calendar year. Any existing trailing year in
//     the subject is replaced so there are no duplicates.
//  2. Locates the "Change log" table, checks the date in the last non-empty
//     data row (DD.MM.YYYY). If that date differs from today, a new row is
//     inserted before the first empty row (or appended if none) with:
//     - incremented major version (e.g. 1.0 → 2.0) taken from last non-empty row
//     - today's date (DD.MM.YYYY)
//     - same Author as the last non-empty row
//     - Description "Release for YYYY"
//  3. Sets (or creates) the custom property "Version" to the version value
//     written in the Change log row.
//
// -dir and -out default to the directory where the executable is located
// when not specified.
//
// Build a standalone Windows executable:
//
//	go build -o changelog_updater.exe tools/changelog_updater.go
//
// Or cross-compile from any OS:
//
//	GOOS=windows GOARCH=amd64 go build -o changelog_updater.exe tools/changelog_updater.go
//
// Usage:
//
//	changelog_updater.exe -dir C:\path\to\docs
//	changelog_updater.exe -dir C:\path\to\docs -out C:\path\to\output
//
// Without -out the originals are overwritten in place.
package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	godocx "github.com/falcomza/go-docx"
)

var trailingYearPattern = regexp.MustCompile(`\s*\b(?:19|20)\d{2}\b\s*$`)

func main() {
	dir := flag.String("dir", "", "Folder containing DOCX files (required)")
	out := flag.String("out", "", "Output folder (omit to overwrite in-place)")

	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, `changelog_updater — batch-update DOCX annual release documents

For every .docx file in -dir:
  1. Sets the document Subject (core property) to "<original subject> YYYY"
     where YYYY is the current calendar year. Any existing trailing year
     in the subject is replaced (e.g. "OSS Check 2024" -> "OSS Check 2026").
  2. Finds the "Change log" table. Reads version and author from the last
     non-empty data row. If its date differs from today, inserts a new row
     before the first empty row (or at the end if no empty rows exist):
       Version      incremented major from last non-empty row (e.g. 1.0 -> 2.0)
       Date         today in DD.MM.YYYY
       Author       copied from the last non-empty row
       Description  "Release for YYYY"
  3. Sets the custom property "Version" (creates it if absent) to the version
     value used in the Change log row.

-dir and -out default to the folder containing the executable when omitted.

Build a Windows executable:
  go build -o changelog_updater.exe tools/changelog_updater.go

Usage:
  changelog_updater.exe -dir <folder> [-out <folder>]

Flags:
`)
		flag.PrintDefaults()
	}

	flag.Parse()

	// Default -dir and -out to the directory containing the executable.
	exeDir := executableDir()
	if *dir == "" {
		*dir = exeDir
	}
	if *out == "" {
		*out = *dir
	}

	entries, err := os.ReadDir(*dir)
	if err != nil {
		log.Fatalf("read dir %s: %v", *dir, err)
	}

	if *out != "" {
		if err := os.MkdirAll(*out, 0o755); err != nil {
			log.Fatalf("create output dir: %v", err)
		}
	}

	now := time.Now()
	today := now.Format("02.01.2006")
	year := strconv.Itoa(now.Year())
	processed, skipped := 0, 0

	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		name := e.Name()
		if !strings.EqualFold(filepath.Ext(name), ".docx") {
			continue
		}

		inputPath := filepath.Join(*dir, name)
		outPath := filepath.Join(*out, name)

		if err := processFile(inputPath, outPath, today, year); err != nil {
			log.Printf("SKIP  %s: %v", name, err)
			skipped++
		} else {
			log.Printf("OK    %s -> %s", name, outPath)
			processed++
		}
	}

	log.Printf("Done: %d processed, %d skipped", processed, skipped)
}

func processFile(inputPath, outPath, today, year string) error {
	u, err := godocx.New(inputPath)
	if err != nil {
		return fmt.Errorf("open: %w", err)
	}
	defer u.Cleanup()

	// 1. Set core properties subject to "<original subject> YYYY".
	props, err := u.GetCoreProperties()
	if err != nil {
		return fmt.Errorf("get core properties: %w", err)
	}
	props.Subject = setTitleYear(props.Subject, year)
	if err := u.SetCoreProperties(*props); err != nil {
		return fmt.Errorf("set core properties: %w", err)
	}

	// 2. Find the Change log table and conditionally insert a row.
	tables, err := u.GetTableText()
	if err != nil {
		return fmt.Errorf("get table text: %w", err)
	}

	tableIdx := findChangeLogTable(tables)
	version := ""
	if tableIdx < 0 {
		log.Printf("      (no Change log table found in %s)", filepath.Base(inputPath))
	} else {
		var verr error
		version, verr = maybeInsertRow(u, tables, tableIdx, today, year)
		if verr != nil {
			return verr
		}
	}

	// 3. Set (or create) the custom property "Version".
	if version != "" {
		if err := setCustomVersion(u, version); err != nil {
			return fmt.Errorf("set custom version: %w", err)
		}
	}

	if err := u.Save(outPath); err != nil {
		return fmt.Errorf("save: %w", err)
	}
	return nil
}

// findChangeLogTable returns the 0-based index of the first table whose first
// row contains a cell with "change" (case-insensitive), or -1 if not found.
func findChangeLogTable(tables [][][]string) int {
	for i, tbl := range tables {
		if len(tbl) == 0 {
			continue
		}
		for _, cell := range tbl[0] {
			if strings.Contains(strings.ToLower(cell), "change") {
				return i
			}
		}
	}
	return -1
}

// maybeInsertRow checks the last non-empty data row of the Change log table
// and inserts a new row (before the first empty row, or at the end) if
// today's date is not already in that row.
// It returns the version string that is current after the operation
// (either the existing version if no row was added, or the incremented one).
func maybeInsertRow(u *godocx.Updater, tables [][][]string, tableIdx int, today, year string) (string, error) {
	tbl := tables[tableIdx]

	// Skip if the table only has a header row (or is empty).
	if len(tbl) < 2 {
		log.Printf("      (Change log table has no data rows, skipping row insert)")
		return "", nil
	}

	// Find the last non-empty data row (skip header at index 0).
	lastNonEmptyIdx := -1
	for i := len(tbl) - 1; i >= 1; i-- {
		if !isEmptyRow(tbl[i]) {
			lastNonEmptyIdx = i
			break
		}
	}
	if lastNonEmptyIdx < 0 {
		log.Printf("      (Change log table has no non-empty data rows, skipping)")
		return "", nil
	}

	lastRow := tbl[lastNonEmptyIdx]
	if len(lastRow) < 2 {
		return "", fmt.Errorf("last non-empty Change log row has fewer than 2 cells")
	}

	// Already up to date — return existing version.
	if strings.TrimSpace(lastRow[1]) == today {
		return strings.TrimSpace(lastRow[0]), nil
	}

	// Build new row values from the last non-empty row.
	newVersion := incrementMajorVersion(strings.TrimSpace(lastRow[0]))
	author := ""
	if len(lastRow) > 2 {
		author = strings.TrimSpace(lastRow[2])
	}
	description := "Release for " + year

	newRow := make([]string, len(lastRow))
	newRow[0] = newVersion
	newRow[1] = today
	if len(newRow) > 2 {
		newRow[2] = author
	}
	if len(newRow) > 3 {
		newRow[3] = description
	}

	// Insert before the first empty row after the last non-empty row,
	// or append at the end if no empty row exists.
	for i := lastNonEmptyIdx + 1; i < len(tbl); i++ {
		if isEmptyRow(tbl[i]) {
			// InsertTableRowBefore uses 1-based indices.
			return newVersion, u.InsertTableRowBefore(tableIdx+1, i+1, newRow)
		}
	}
	return newVersion, u.AppendTableRow(tableIdx+1, newRow)
}

// isEmptyRow returns true when every cell in the row is blank after trimming.
func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

// setCustomVersion upserts the "Version" custom property.
func setCustomVersion(u *godocx.Updater, version string) error {
	existing, err := u.GetCustomProperties()
	if err != nil {
		return err
	}
	// Replace if exists, otherwise append.
	updated := make([]godocx.CustomProperty, 0, len(existing)+1)
	found := false
	for _, p := range existing {
		if strings.EqualFold(p.Name, "Version") {
			p.Value = version
			p.Type = "lpwstr"
			found = true
		}
		updated = append(updated, p)
	}
	if !found {
		updated = append(updated, godocx.CustomProperty{Name: "Version", Value: version, Type: "lpwstr"})
	}
	return u.SetCustomProperties(updated)
}

// executableDir returns the directory that contains the running executable.
// Falls back to the current working directory on error.
func executableDir() string {
	exe, err := os.Executable()
	if err != nil {
		wd, _ := os.Getwd()
		return wd
	}
	return filepath.Dir(exe)
}

// year, or year appended if none was present.
// E.g. setTitleYear("OSS 2024", "2026") == "OSS 2026"
//
//	setTitleYear("OSS Check", "2026") == "OSS Check 2026"
func setTitleYear(title, year string) string {
	base := strings.TrimSpace(trailingYearPattern.ReplaceAllString(title, ""))
	if base == "" {
		return year
	}
	return base + " " + year
}

// incrementMajorVersion parses a version string like "1.0" or "2.3" and
// returns the version with the major component incremented by 1 and the
// minor component reset to 0 (e.g. 1.0 → 2.0, 3.5 → 4.0).
// If parsing fails, the original string is returned unchanged.
func incrementMajorVersion(v string) string {
	parts := strings.SplitN(v, ".", 2)
	if len(parts) != 2 {
		return v
	}
	major, err := strconv.Atoi(strings.TrimSpace(parts[0]))
	if err != nil {
		return v
	}
	return fmt.Sprintf("%d.0", major+1)
}

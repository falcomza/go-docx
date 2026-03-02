package godocx

import (
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// xmlEscapeReplacer normalises the numeric entity refs that xml.EscapeText emits
// for quote characters into the named-entity forms expected by the Word XML
// serialiser and the existing test suite.
var xmlEscapeReplacer = strings.NewReplacer("&#34;", "&quot;", "&#39;", "&apos;")

// xmlEscape escapes XML special characters using the stdlib XML encoder.
// This handles all edge cases including carriage-returns (&#xD;) and invalid
// UTF-8 sequences, which a naive strings.ReplaceAll approach would miss.
// Named entities (&quot;, &apos;) are used instead of numeric ones.
func xmlEscape(s string) string {
	var b strings.Builder
	if err := xml.EscapeText(&b, []byte(s)); err != nil {
		// Fallback to manual escaping if the stdlib encoder fails.
		return strings.NewReplacer(
			"&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;", "'", "&apos;",
		).Replace(s)
	}
	return xmlEscapeReplacer.Replace(b.String())
}

// formatFloat formats a float64 for XML output.
// Removes trailing zeros and unnecessary decimal points.
func formatFloat(f float64) string {
	return strconv.FormatFloat(f, 'f', -1, 64)
}

// columnLetters converts a column number (1-based) to Excel column letters.
// Examples: 1->A, 2->B, 26->Z, 27->AA, 28->AB
func columnLetters(n int) string {
	if n <= 0 {
		return "A"
	}
	// Write digits right-to-left into a fixed-size buffer to avoid
	// the O(n²) repeated slice prepend of the naive approach.
	var buf [8]byte
	i := len(buf)
	for n > 0 {
		n--
		i--
		buf[i] = byte('A' + n%26)
		n /= 26
	}
	return string(buf[i:])
}

// cellRef generates an Excel cell reference from column and row numbers.
// Both col and row are 1-based. Example: cellRef(1, 1) -> "A1"
func cellRef(col, row int) string {
	return columnLetters(col) + strconv.Itoa(row)
}

// normalizeHexColor normalizes a hex color code for use in Office Open XML.
// Accepts colors with or without '#' prefix. Returns empty string if invalid.
// Examples: "#FF0000" -> "FF0000", "ff0000" -> "FF0000"
func normalizeHexColor(color string) string {
	c := strings.TrimSpace(color)
	if c == "" {
		return ""
	}
	c = strings.TrimPrefix(c, "#")
	if len(c) != 6 {
		return ""
	}
	for _, ch := range c {
		if !(ch >= '0' && ch <= '9' || ch >= 'a' && ch <= 'f' || ch >= 'A' && ch <= 'F') {
			return ""
		}
	}
	return strings.ToUpper(c)
}

// getNextDocPrId finds the next available docPr ID in the document.
func (u *Updater) getNextDocPrId() (int, error) {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document: %w", err)
	}

	matches := docPrIDPattern.FindAllStringSubmatch(string(raw), -1)

	maxId := 0
	for _, match := range matches {
		if len(match) > 1 {
			id, err := strconv.Atoi(match[1])
			if err != nil {
				continue
			}
			if id > maxId {
				maxId = id
			}
		}
	}

	return maxId + 1, nil
}

// getNextDocumentRelId finds the next available relationship ID in document.xml.rels.
func (u *Updater) getNextDocumentRelId() (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	return getNextRelIDFromFile(relsPath)
}

// atomicWriteFile writes data to path atomically using a write-then-rename
// strategy. This prevents a partially-written (corrupted) file being visible
// to readers if the process crashes mid-write on a local filesystem.
func atomicWriteFile(path string, data []byte, perm os.FileMode) error {
	dir := filepath.Dir(path)
	tmp, err := os.CreateTemp(dir, ".docx-write-*")
	if err != nil {
		return fmt.Errorf("create temp file: %w", err)
	}
	tmpName := tmp.Name()
	if _, err := tmp.Write(data); err != nil {
		tmp.Close()
		os.Remove(tmpName)
		return fmt.Errorf("write temp file: %w", err)
	}
	if err := tmp.Close(); err != nil {
		os.Remove(tmpName)
		return fmt.Errorf("close temp file: %w", err)
	}
	if err := os.Chmod(tmpName, perm); err != nil {
		os.Remove(tmpName)
		return fmt.Errorf("chmod temp file: %w", err)
	}
	// On Windows, os.Rename fails with "Access is denied" when the destination
	// already exists.  Remove it first; this sacrifices strict atomicity on
	// Windows but is safe because we are always writing into a private temp dir.
	_ = os.Remove(path)
	if err := os.Rename(tmpName, path); err != nil {
		os.Remove(tmpName)
		return fmt.Errorf("rename temp file: %w", err)
	}
	return nil
}

// getNextRelIDFromFile finds the next available relationship ID in a .rels file.
func getNextRelIDFromFile(relsPath string) (string, error) {
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read rels file %s: %w", relsPath, err)
	}

	var rels relationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return "", fmt.Errorf("parse rels file %s: %w", relsPath, err)
	}

	maxId := 0
	for _, rel := range rels.Relationships {
		if matches := relIDPattern.FindStringSubmatch(rel.ID); matches != nil {
			id, err := strconv.Atoi(matches[1])
			if err != nil {
				continue
			}
			if id > maxId {
				maxId = id
			}
		}
	}

	return fmt.Sprintf("rId%d", maxId+1), nil
}

package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
)

// DeleteOptions defines options for content deletion
type DeleteOptions struct {
	// MatchCase determines if search is case-sensitive
	MatchCase bool

	// WholeWord only matches whole words
	WholeWord bool
}

// DefaultDeleteOptions returns delete options with sensible defaults
func DefaultDeleteOptions() DeleteOptions {
	return DeleteOptions{
		MatchCase: false,
		WholeWord: false,
	}
}

// DeleteParagraphs removes paragraphs matching the specified text.
// Returns the number of paragraphs deleted.
func (u *Updater) DeleteParagraphs(text string, opts DeleteOptions) (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}
	if text == "" {
		return 0, fmt.Errorf("text cannot be empty")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document.xml: %w", err)
	}

	updated, count, err := deleteParagraphsContaining(raw, text, opts)
	if err != nil {
		return count, fmt.Errorf("delete paragraphs: %w", err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return count, fmt.Errorf("write document.xml: %w", err)
	}

	return count, nil
}

// DeleteTable removes a table by index (1-based).
func (u *Updater) DeleteTable(tableIndex int) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if tableIndex < 1 {
		return fmt.Errorf("table index must be >= 1")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := deleteNthTable(raw, tableIndex)
	if err != nil {
		return fmt.Errorf("delete table %d: %w", tableIndex, err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// DeleteImage removes an image by index (1-based).
// Note: This removes the image from the document but does not delete the media file.
func (u *Updater) DeleteImage(imageIndex int) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if imageIndex < 1 {
		return fmt.Errorf("image index must be >= 1")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := deleteNthImage(raw, imageIndex)
	if err != nil {
		return fmt.Errorf("delete image %d: %w", imageIndex, err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// DeleteChart removes a chart by index (1-based).
func (u *Updater) DeleteChart(chartIndex int) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if chartIndex < 1 {
		return fmt.Errorf("chart index must be >= 1")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := deleteNthChart(raw, chartIndex)
	if err != nil {
		return fmt.Errorf("delete chart %d: %w", chartIndex, err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// deleteParagraphsContaining removes paragraphs that contain the specified text
func deleteParagraphsContaining(raw []byte, text string, opts DeleteOptions) ([]byte, int, error) {
	// Build search pattern
	var pattern *regexp.Regexp
	if opts.WholeWord {
		wordBoundary := `\b` + regexp.QuoteMeta(text) + `\b`
		if !opts.MatchCase {
			pattern = regexp.MustCompile("(?i)" + wordBoundary)
		} else {
			pattern = regexp.MustCompile(wordBoundary)
		}
	} else {
		if !opts.MatchCase {
			pattern = regexp.MustCompile("(?i)" + regexp.QuoteMeta(text))
		} else {
			pattern = regexp.MustCompile(regexp.QuoteMeta(text))
		}
	}

	// Find all paragraphs
	paras := extractParaPattern.FindAllIndex(raw, -1)

	count := 0
	var result bytes.Buffer
	lastEnd := 0

	for _, paraIdx := range paras {
		paraContent := raw[paraIdx[0]:paraIdx[1]]
		paraText := extractParagraphPlainText(paraContent)

		if pattern.MatchString(paraText) {
			// Flush content before this paragraph, then skip it
			result.Write(raw[lastEnd:paraIdx[0]])
			lastEnd = paraIdx[1]
			count++
			continue
		}

		// Keep this paragraph
		result.Write(raw[lastEnd:paraIdx[1]])
		lastEnd = paraIdx[1]
	}

	// Write remaining content after last paragraph
	result.Write(raw[lastEnd:])

	return result.Bytes(), count, nil
}

// deleteNthTable removes the Nth table from the document
func deleteNthTable(raw []byte, n int) ([]byte, error) {
	// Find all tables
	tables := extractTablePattern.FindAllIndex(raw, -1)

	if n > len(tables) {
		return nil, fmt.Errorf("table %d not found (document has %d tables)", n, len(tables))
	}

	tableIdx := tables[n-1]

	// Build result without this table
	var result bytes.Buffer
	result.Write(raw[:tableIdx[0]])
	result.Write(raw[tableIdx[1]:])

	return result.Bytes(), nil
}

// deleteNthImage removes the Nth image (drawing with blip) from the document
func deleteNthImage(raw []byte, n int) ([]byte, error) {
	// Find all image drawings (wp:inline with a:blip)
	// Pattern matches the paragraph containing an image
	imagePattern := regexp.MustCompile(`(?s)<w:p[^>]*>.*?<wp:inline.*?<a:blip[^>]*r:embed="[^"]*"[^>]*>.*?</wp:inline>.*?</w:p>`)
	images := imagePattern.FindAllIndex(raw, -1)

	if n > len(images) {
		return nil, fmt.Errorf("image %d not found (document has %d images)", n, len(images))
	}

	imgIdx := images[n-1]

	// Build result without this image
	var result bytes.Buffer
	result.Write(raw[:imgIdx[0]])
	result.Write(raw[imgIdx[1]:])

	return result.Bytes(), nil
}

// deleteNthChart removes the Nth chart from the document
func deleteNthChart(raw []byte, n int) ([]byte, error) {
	// Find all chart drawings (wp:inline with c:chart)
	chartPattern := regexp.MustCompile(`(?s)<w:p[^>]*>.*?<wp:inline.*?<c:chart[^>]*r:id="[^"]*"[^>]*>.*?</wp:inline>.*?</w:p>`)
	charts := chartPattern.FindAllIndex(raw, -1)

	if n > len(charts) {
		return nil, fmt.Errorf("chart %d not found (document has %d charts)", n, len(charts))
	}

	chartIdx := charts[n-1]

	// Build result without this chart
	var result bytes.Buffer
	result.Write(raw[:chartIdx[0]])
	result.Write(raw[chartIdx[1]:])

	return result.Bytes(), nil
}

// GetTableCount returns the number of tables in the document
func (u *Updater) GetTableCount() (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document.xml: %w", err)
	}

	tablePattern := regexp.MustCompile(`(?s)<w:tbl>`)
	tables := tablePattern.FindAllIndex(raw, -1)

	return len(tables), nil
}

// GetParagraphCount returns the number of paragraphs in the document
func (u *Updater) GetParagraphCount() (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document.xml: %w", err)
	}

	paraPattern := regexp.MustCompile(`(?s)<w:p[^>]*>`)
	paras := paraPattern.FindAllIndex(raw, -1)

	return len(paras), nil
}

// GetImageCount returns the number of images in the document
func (u *Updater) GetImageCount() (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document.xml: %w", err)
	}

	// Count images by counting blip elements
	blipPattern := regexp.MustCompile(`<a:blip[^>]*r:embed="[^"]*"[^>]*>`)
	blips := blipPattern.FindAllIndex(raw, -1)

	return len(blips), nil
}


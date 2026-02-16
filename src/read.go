package docxupdater

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// TextMatch represents a text match with context
type TextMatch struct {
	Text      string // The matched text
	Paragraph int    // Paragraph index (0-based)
	Position  int    // Character position in document
	Before    string // Context before match (up to 50 chars)
	After     string // Context after match (up to 50 chars)
}

// FindOptions defines options for text search
type FindOptions struct {
	// MatchCase determines if search is case-sensitive
	MatchCase bool

	// WholeWord only matches whole words
	WholeWord bool

	// UseRegex treats the search string as a regular expression
	UseRegex bool

	// MaxResults limits the number of results (0 for unlimited)
	MaxResults int

	// InParagraphs enables search in paragraphs
	InParagraphs bool

	// InTables enables search in tables
	InTables bool

	// InHeaders enables search in headers
	InHeaders bool

	// InFooters enables search in footers
	InFooters bool
}

// DefaultFindOptions returns find options with sensible defaults
func DefaultFindOptions() FindOptions {
	return FindOptions{
		MatchCase:    false,
		WholeWord:    false,
		UseRegex:     false,
		MaxResults:   0, // unlimited
		InParagraphs: true,
		InTables:     true,
		InHeaders:    false,
		InFooters:    false,
	}
}

// GetText extracts all text from the document body
func (u *Updater) GetText() (string, error) {
	if u == nil {
		return "", fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return "", NewXMLParseError("document.xml", err)
	}

	return u.extractTextFromXML(raw), nil
}

// GetParagraphText extracts text from all paragraphs
func (u *Updater) GetParagraphText() ([]string, error) {
	if u == nil {
		return nil, fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return nil, NewXMLParseError("document.xml", err)
	}

	return u.extractParagraphsFromXML(raw), nil
}

// GetTableText extracts text from all tables
// Returns a 2D slice where each element represents a table, containing rows of cells
func (u *Updater) GetTableText() ([][][]string, error) {
	if u == nil {
		return nil, fmt.Errorf("updater is nil")
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return nil, NewXMLParseError("document.xml", err)
	}

	return u.extractTablesFromXML(raw), nil
}

// FindText finds all occurrences of text in the document
func (u *Updater) FindText(pattern string, opts FindOptions) ([]TextMatch, error) {
	if u == nil {
		return nil, fmt.Errorf("updater is nil")
	}
	if pattern == "" {
		return nil, NewValidationError("pattern", "search pattern cannot be empty")
	}

	var matches []TextMatch
	var searchPattern *regexp.Regexp
	var err error

	// Compile search pattern
	if opts.UseRegex {
		if !opts.MatchCase {
			pattern = "(?i)" + pattern
		}
		searchPattern, err = regexp.Compile(pattern)
		if err != nil {
			return nil, NewInvalidRegexError(pattern, err)
		}
	} else {
		// Escape regex metachars for literal search
		escapedPattern := regexp.QuoteMeta(pattern)
		if opts.WholeWord {
			escapedPattern = `\b` + escapedPattern + `\b`
		}
		if !opts.MatchCase {
			escapedPattern = "(?i)" + escapedPattern
		}
		searchPattern = regexp.MustCompile(escapedPattern)
	}

	// Search in document body
	if opts.InParagraphs || opts.InTables {
		docPath := filepath.Join(u.tempDir, "word", "document.xml")
		raw, err := os.ReadFile(docPath)
		if err != nil {
			return nil, NewXMLParseError("document.xml", err)
		}

		docMatches := u.findInXML(raw, searchPattern, opts.MaxResults-len(matches))
		matches = append(matches, docMatches...)
	}

	// Search in headers
	if opts.InHeaders && (opts.MaxResults == 0 || len(matches) < opts.MaxResults) {
		headerFiles, _ := filepath.Glob(filepath.Join(u.tempDir, "word", "header*.xml"))
		for _, headerPath := range headerFiles {
			raw, err := os.ReadFile(headerPath)
			if err != nil {
				continue
			}
			headerMatches := u.findInXML(raw, searchPattern, opts.MaxResults-len(matches))
			matches = append(matches, headerMatches...)
			if opts.MaxResults > 0 && len(matches) >= opts.MaxResults {
				break
			}
		}
	}

	// Search in footers
	if opts.InFooters && (opts.MaxResults == 0 || len(matches) < opts.MaxResults) {
		footerFiles, _ := filepath.Glob(filepath.Join(u.tempDir, "word", "footer*.xml"))
		for _, footerPath := range footerFiles {
			raw, err := os.ReadFile(footerPath)
			if err != nil {
				continue
			}
			footerMatches := u.findInXML(raw, searchPattern, opts.MaxResults-len(matches))
			matches = append(matches, footerMatches...)
			if opts.MaxResults > 0 && len(matches) >= opts.MaxResults {
				break
			}
		}
	}

	return matches, nil
}

// extractTextFromXML extracts all visible text from XML content
func (u *Updater) extractTextFromXML(raw []byte) string {
	var result strings.Builder

	// Extract text from <w:t> elements
	textPattern := regexp.MustCompile(`<w:t[^>]*>(.*?)</w:t>`)
	matches := textPattern.FindAllSubmatch(raw, -1)

	for _, match := range matches {
		if len(match) > 1 {
			text := string(match[1])
			text = unescapeXML(text)
			result.WriteString(text)
		}
	}

	return result.String()
}

// extractParagraphsFromXML extracts text from each paragraph
func (u *Updater) extractParagraphsFromXML(raw []byte) []string {
	var paragraphs []string

	// Find all <w:p> elements
	paraPattern := regexp.MustCompile(`<w:p[^>]*>.*?</w:p>`)
	paraMatches := paraPattern.FindAll(raw, -1)

	for _, paraByte := range paraMatches {
		// Extract text from this paragraph
		text := u.extractTextFromXML(paraByte)
		if text != "" {
			paragraphs = append(paragraphs, text)
		}
	}

	return paragraphs
}

// extractTablesFromXML extracts text from all tables
func (u *Updater) extractTablesFromXML(raw []byte) [][][]string {
	var tables [][][]string

	// Find all <w:tbl> elements
	tablePattern := regexp.MustCompile(`<w:tbl>.*?</w:tbl>`)
	tableMatches := tablePattern.FindAll(raw, -1)

	for _, tableBytes := range tableMatches {
		table := u.extractTableData(tableBytes)
		if len(table) > 0 {
			tables = append(tables, table)
		}
	}

	return tables
}

// extractTableData extracts rows and cells from a table
func (u *Updater) extractTableData(tableXML []byte) [][]string {
	var rows [][]string

	// Find all <w:tr> (table row) elements
	rowPattern := regexp.MustCompile(`<w:tr[^>]*>.*?</w:tr>`)
	rowMatches := rowPattern.FindAll(tableXML, -1)

	for _, rowBytes := range rowMatches {
		var cells []string

		// Find all <w:tc> (table cell) elements
		cellPattern := regexp.MustCompile(`<w:tc>.*?</w:tc>`)
		cellMatches := cellPattern.FindAll(rowBytes, -1)

		for _, cellBytes := range cellMatches {
			cellText := u.extractTextFromXML(cellBytes)
			cells = append(cells, cellText)
		}

		if len(cells) > 0 {
			rows = append(rows, cells)
		}
	}

	return rows
}

// findInXML finds all matches of the pattern in XML content
func (u *Updater) findInXML(raw []byte, pattern *regexp.Regexp, maxResults int) []TextMatch {
	var matches []TextMatch

	// Extract full text for searching
	fullText := u.extractTextFromXML(raw)

	// Find paragraphs for indexing
	paragraphs := u.extractParagraphsFromXML(raw)

	// Find all matches in the full text
	indices := pattern.FindAllStringIndex(fullText, -1)

	for _, idx := range indices {
		if maxResults > 0 && len(matches) >= maxResults {
			break
		}

		matchText := fullText[idx[0]:idx[1]]

		// Determine which paragraph this match belongs to
		paraIndex := u.findParagraphIndex(fullText, idx[0], paragraphs)

		// Extract context (50 chars before and after)
		contextBefore := ""
		contextAfter := ""

		beforeStart := idx[0] - 50
		if beforeStart < 0 {
			beforeStart = 0
		}
		contextBefore = fullText[beforeStart:idx[0]]

		afterEnd := idx[1] + 50
		if afterEnd > len(fullText) {
			afterEnd = len(fullText)
		}
		contextAfter = fullText[idx[1]:afterEnd]

		matches = append(matches, TextMatch{
			Text:      matchText,
			Paragraph: paraIndex,
			Position:  idx[0],
			Before:    contextBefore,
			After:     contextAfter,
		})
	}

	return matches
}

// findParagraphIndex determines which paragraph contains the given position
func (u *Updater) findParagraphIndex(fullText string, position int, paragraphs []string) int {
	currentPos := 0
	for i, para := range paragraphs {
		paraLen := len(para)
		if position >= currentPos && position < currentPos+paraLen {
			return i
		}
		currentPos += paraLen
	}
	return -1
}

// unescapeXML unescapes XML entities
func unescapeXML(s string) string {
	s = strings.ReplaceAll(s, "&amp;", "&")
	s = strings.ReplaceAll(s, "&lt;", "<")
	s = strings.ReplaceAll(s, "&gt;", ">")
	s = strings.ReplaceAll(s, "&quot;", "\"")
	s = strings.ReplaceAll(s, "&apos;", "'")
	return s
}

package godocx

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// ReplaceOptions defines options for text replacement
type ReplaceOptions struct {
	// MatchCase determines if replacement is case-sensitive
	MatchCase bool

	// WholeWord only replaces whole word matches
	WholeWord bool

	// InParagraphs enables replacement in paragraph text
	InParagraphs bool

	// InTables enables replacement in table cells
	InTables bool

	// InHeaders enables replacement in headers
	InHeaders bool

	// InFooters enables replacement in footers
	InFooters bool

	// MaxReplacements limits the number of replacements (0 for unlimited)
	MaxReplacements int
}

// DefaultReplaceOptions returns replace options with sensible defaults
func DefaultReplaceOptions() ReplaceOptions {
	return ReplaceOptions{
		MatchCase:       false,
		WholeWord:       false,
		InParagraphs:    true,
		InTables:        true,
		InHeaders:       false,
		InFooters:       false,
		MaxReplacements: 0, // unlimited
	}
}

// ReplaceText replaces all occurrences of old text with new text
// Returns the number of replacements made
func (u *Updater) ReplaceText(old, new string, opts ReplaceOptions) (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}
	if old == "" {
		return 0, NewValidationError("old", "old text cannot be empty")
	}

	count := 0

	// Replace in document body (paragraphs and tables)
	if opts.InParagraphs || opts.InTables {
		docPath := filepath.Join(u.tempDir, "word", "document.xml")
		_, err := u.replaceInFile(docPath, old, new, opts, &count)
		if err != nil {
			return count, fmt.Errorf("replace in document: %w", err)
		}
	}

	// Replace in headers
	if opts.InHeaders {
		headerFiles, _ := filepath.Glob(filepath.Join(u.tempDir, "word", "header*.xml"))
		for _, headerPath := range headerFiles {
			_, err := u.replaceInFile(headerPath, old, new, opts, &count)
			if err != nil {
				return count, fmt.Errorf("replace in header: %w", err)
			}
		}
	}

	// Replace in footers
	if opts.InFooters {
		footerFiles, _ := filepath.Glob(filepath.Join(u.tempDir, "word", "footer*.xml"))
		for _, footerPath := range footerFiles {
			_, err := u.replaceInFile(footerPath, old, new, opts, &count)
			if err != nil {
				return count, fmt.Errorf("replace in footer: %w", err)
			}
		}
	}

	return count, nil
}

// ReplaceTextRegex replaces text matching a regular expression pattern
// Returns the number of replacements made
func (u *Updater) ReplaceTextRegex(pattern *regexp.Regexp, replacement string, opts ReplaceOptions) (int, error) {
	if u == nil {
		return 0, fmt.Errorf("updater is nil")
	}
	if pattern == nil {
		return 0, NewValidationError("pattern", "regex pattern cannot be nil")
	}

	count := 0

	// Replace in document body
	if opts.InParagraphs || opts.InTables {
		docPath := filepath.Join(u.tempDir, "word", "document.xml")
		_, err := u.replaceRegexInFile(docPath, pattern, replacement, opts, &count)
		if err != nil {
			return count, fmt.Errorf("replace in document: %w", err)
		}
	}

	// Replace in headers
	if opts.InHeaders {
		headerFiles, _ := filepath.Glob(filepath.Join(u.tempDir, "word", "header*.xml"))
		for _, headerPath := range headerFiles {
			_, err := u.replaceRegexInFile(headerPath, pattern, replacement, opts, &count)
			if err != nil {
				return count, fmt.Errorf("replace in header: %w", err)
			}
		}
	}

	// Replace in footers
	if opts.InFooters {
		footerFiles, _ := filepath.Glob(filepath.Join(u.tempDir, "word", "footer*.xml"))
		for _, footerPath := range footerFiles {
			_, err := u.replaceRegexInFile(footerPath, pattern, replacement, opts, &count)
			if err != nil {
				return count, fmt.Errorf("replace in footer: %w", err)
			}
		}
	}

	return count, nil
}

// normalizeRunsInXML merges consecutive <w:r> elements that share the same
// <w:rPr> (run properties) within each paragraph. This pre-processing pass
// ensures template placeholders like {{FIELD}} that Word fragmented across
// multiple text runs are reconstituted into a single run before substitution.
// The operation is idempotent and preserves all paragraph/character formatting.
func normalizeRunsInXML(raw []byte) []byte {
	return extractParaPattern.ReplaceAllFunc(raw, mergeCompatibleRuns)
}

// mergeCompatibleRuns consolidates consecutive <w:r> elements with identical
// run properties within a single paragraph. Returns the paragraph unchanged if
// no merging opportunities are found.
func mergeCompatibleRuns(para []byte) []byte {
	s := string(para)

	type runSpan struct {
		start, end   int
		rpr          string // empty string if the run has no <w:rPr>
		text         string // decoded text from all <w:t> elements in the run
		needPreserve bool   // whether xml:space="preserve" is required
	}

	// Locate every <w:r> element and record its position and content.
	rawLocs := runBlockPattern.FindAllStringIndex(s, -1)
	if len(rawLocs) <= 1 {
		return para // nothing to merge
	}

	spans := make([]runSpan, 0, len(rawLocs))
	for _, loc := range rawLocs {
		runStr := s[loc[0]:loc[1]]
		rpr := runRprPattern.FindString(runStr)

		// Collect and decode text from any <w:t> elements inside this run.
		var sb strings.Builder
		needPreserve := false
		for _, tm := range extractTextPattern.FindAllStringSubmatch(runStr, -1) {
			sb.WriteString(xmlUnescape(tm[1]))
			if strings.Contains(tm[0], `xml:space="preserve"`) {
				needPreserve = true
			}
		}
		text := sb.String()
		if !needPreserve && (strings.HasPrefix(text, " ") || strings.HasSuffix(text, " ")) {
			needPreserve = true
		}

		spans = append(spans, runSpan{
			start:        loc[0],
			end:          loc[1],
			rpr:          rpr,
			text:         text,
			needPreserve: needPreserve,
		})
	}

	// Group consecutive spans that share identical rPr.
	type mergedSpan struct {
		start, end   int
		rpr          string
		text         string
		needPreserve bool
	}

	merged := make([]mergedSpan, 0, len(spans))
	cur := mergedSpan{
		start:        spans[0].start,
		end:          spans[0].end,
		rpr:          spans[0].rpr,
		text:         spans[0].text,
		needPreserve: spans[0].needPreserve,
	}
	for i := 1; i < len(spans); i++ {
		// Only merge if (a) same run properties AND (b) nothing between the two
		// runs except optional whitespace. Any XML tag between them (e.g.,
		// </w:hyperlink>, <w:bookmarkEnd/>, </w:ins>) means the runs are in
		// different structural contexts and must not be collapsed.
		between := s[cur.end:spans[i].start]
		if spans[i].rpr == cur.rpr && !strings.ContainsAny(between, "<>") {
			// Same run properties with no intervening markup — safe to merge.
			cur.end = spans[i].end
			cur.text += spans[i].text
			if spans[i].needPreserve {
				cur.needPreserve = true
			}
		} else {
			merged = append(merged, cur)
			cur = mergedSpan{
				start:        spans[i].start,
				end:          spans[i].end,
				rpr:          spans[i].rpr,
				text:         spans[i].text,
				needPreserve: spans[i].needPreserve,
			}
		}
	}
	merged = append(merged, cur)

	// No merging took place — return original bytes.
	if len(merged) == len(spans) {
		return para
	}

	// Rebuild the paragraph, replacing each merged multi-run span with a
	// single canonical <w:r> that carries the combined text.
	var out strings.Builder
	out.Grow(len(s))
	pos := 0
	for _, m := range merged {
		out.WriteString(s[pos:m.start])
		out.WriteString("<w:r>")
		if m.rpr != "" {
			out.WriteString(m.rpr)
		}
		if m.needPreserve {
			out.WriteString(`<w:t xml:space="preserve">`)
		} else {
			out.WriteString("<w:t>")
		}
		out.WriteString(xmlEscape(m.text))
		out.WriteString("</w:t></w:r>")
		pos = m.end
	}
	out.WriteString(s[pos:])

	return []byte(out.String())
}

// replaceInFile replaces text in a single XML file.
// It normalizes split runs first so that placeholders fragmented across
// multiple <w:r> elements (common in template headers/footers) are found.
func (u *Updater) replaceInFile(path, old, new string, opts ReplaceOptions, count *int) (int, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return 0, err
	}

	// Merge consecutive compatible runs so split placeholders are visible.
	raw = normalizeRunsInXML(raw)

	updated, replaced := u.replaceTextInXML(raw, old, new, opts, count)
	if replaced > 0 {
		if err := os.WriteFile(path, updated, 0o644); err != nil {
			return 0, err
		}
	}

	return replaced, nil
}

// replaceRegexInFile replaces text matching regex in a single XML file.
// It normalizes split runs first so that patterns spanning multiple <w:r>
// elements (common in template headers/footers) are matched correctly.
func (u *Updater) replaceRegexInFile(path string, pattern *regexp.Regexp, replacement string, opts ReplaceOptions, count *int) (int, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return 0, err
	}

	// Merge consecutive compatible runs so split patterns are visible.
	raw = normalizeRunsInXML(raw)

	updated, replaced := u.replaceRegexInXML(raw, pattern, replacement, opts, count)
	if replaced > 0 {
		if err := os.WriteFile(path, updated, 0o644); err != nil {
			return 0, err
		}
	}

	return replaced, nil
}

// replaceTextInXML performs the actual text replacement in XML content
func (u *Updater) replaceTextInXML(raw []byte, old, new string, opts ReplaceOptions, count *int) ([]byte, int) {
	content := string(raw)
	replaced := 0
	escapedOld := regexp.QuoteMeta(old)
	wordRe := (*regexp.Regexp)(nil)
	caseInsensitiveRe := (*regexp.Regexp)(nil)
	if opts.WholeWord {
		wordPattern := fmt.Sprintf(`\b%s\b`, escapedOld)
		if !opts.MatchCase {
			wordPattern = `(?i)` + wordPattern
		}
		wordRe = regexp.MustCompile(wordPattern)
	} else if !opts.MatchCase {
		caseInsensitiveRe = regexp.MustCompile(`(?i)` + escapedOld)
	}

	// Extract text runs (<w:t> elements) and replace within them
	// Use word boundary \b or explicit space/> to avoid matching <w:tabs>, <w:tbl>, <w:tc>, etc.
	content = textRunPattern.ReplaceAllStringFunc(content, func(match string) string {
		// Check if we've hit the max replacements
		if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
			return match
		}

		// Extract the text content
		// Match the full element: <w:t> or <w:t xml:space="preserve">text</w:t>
		matches := textContentPattern.FindStringSubmatch(match)
		if len(matches) < 2 {
			return match
		}

		rawText := matches[1]
		text := xmlUnescape(rawText)
		var replacedText string

		if opts.WholeWord {
			replacedText = wordRe.ReplaceAllStringFunc(text, func(m string) string {
				if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
					return m
				}
				*count++
				replaced++
				return new
			})
		} else if opts.MatchCase {
			// Case-sensitive simple replacement
			replacedText = strings.ReplaceAll(text, old, new)
			occurrences := strings.Count(text, old)
			if occurrences > 0 {
				if opts.MaxReplacements > 0 {
					limit := opts.MaxReplacements - *count
					if occurrences > limit {
						occurrences = limit
					}
				}
				*count += occurrences
				replaced += occurrences
			}
		} else {
			// Case-insensitive replacement
			replacedText = caseInsensitiveRe.ReplaceAllStringFunc(text, func(m string) string {
				if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
					return m
				}
				*count++
				replaced++
				return new
			})
		}

		if replacedText != text {
			return strings.Replace(match, rawText, xmlEscape(replacedText), 1)
		}
		return match
	})

	return []byte(content), replaced
}

// replaceRegexInXML performs regex replacement in XML content
func (u *Updater) replaceRegexInXML(raw []byte, pattern *regexp.Regexp, replacement string, opts ReplaceOptions, count *int) ([]byte, int) {
	content := string(raw)
	replaced := 0

	// Extract text runs (<w:t> elements) and replace within them
	// Use word boundary or explicit space/> to avoid matching <w:tabs>, <w:tbl>, <w:tc>, etc.
	content = textRunPattern.ReplaceAllStringFunc(content, func(match string) string {
		// Check if we've hit the max replacements
		if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
			return match
		}

		// Extract the text content
		// Match the full element: <w:t> or <w:t xml:space="preserve">text</w:t>
		matches := textContentPattern.FindStringSubmatch(match)
		if len(matches) < 2 {
			return match
		}
		rawText := matches[1]
		text := xmlUnescape(rawText)
		replacedText := pattern.ReplaceAllStringFunc(text, func(m string) string {
			if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
				return m
			}
			*count++
			replaced++
			return replacement
		})

		if replacedText != text {
			return strings.Replace(match, rawText, xmlEscape(replacedText), 1)
		}
		return match
	})

	return []byte(content), replaced
}

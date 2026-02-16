package docxupdater

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

// replaceInFile replaces text in a single XML file
func (u *Updater) replaceInFile(path, old, new string, opts ReplaceOptions, count *int) (int, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return 0, err
	}

	updated, replaced := u.replaceTextInXML(raw, old, new, opts, count)
	if replaced > 0 {
		if err := os.WriteFile(path, updated, 0o644); err != nil {
			return 0, err
		}
	}

	return replaced, nil
}

// replaceRegexInFile replaces text matching regex in a single XML file
func (u *Updater) replaceRegexInFile(path string, pattern *regexp.Regexp, replacement string, opts ReplaceOptions, count *int) (int, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return 0, err
	}

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

	// Extract text runs (<w:t> elements) and replace within them
	textPattern := regexp.MustCompile(`<w:t[^>]*>(.*?)</w:t>`)

	content = textPattern.ReplaceAllStringFunc(content, func(match string) string {
		// Check if we've hit the max replacements
		if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
			return match
		}

		// Extract the text content
		textContentPattern := regexp.MustCompile(`<w:t[^>]*>(.*?)</w:t>`)
		matches := textContentPattern.FindStringSubmatch(match)
		if len(matches) < 2 {
			return match
		}

		text := matches[1]
		var replacedText string

		if opts.WholeWord {
			// Use word boundary for whole word matching
			wordPattern := fmt.Sprintf(`\b%s\b`, regexp.QuoteMeta(old))
			if !opts.MatchCase {
				wordPattern = `(?i)` + wordPattern
			}
			re := regexp.MustCompile(wordPattern)
			replacedText = re.ReplaceAllStringFunc(text, func(m string) string {
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
			re := regexp.MustCompile(`(?i)` + regexp.QuoteMeta(old))
			replacedText = re.ReplaceAllStringFunc(text, func(m string) string {
				if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
					return m
				}
				*count++
				replaced++
				return new
			})
		}

		if replacedText != text {
			return strings.Replace(match, text, escapeXML(replacedText), 1)
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
	textPattern := regexp.MustCompile(`<w:t[^>]*>(.*?)</w:t>`)

	content = textPattern.ReplaceAllStringFunc(content, func(match string) string {
		// Check if we've hit the max replacements
		if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
			return match
		}

		// Extract the text content
		textContentPattern := regexp.MustCompile(`<w:t[^>]*>(.*?)</w:t>`)
		matches := textContentPattern.FindStringSubmatch(match)
		if len(matches) < 2 {
			return match
		}

		text := matches[1]
		replacedText := pattern.ReplaceAllStringFunc(text, func(m string) string {
			if opts.MaxReplacements > 0 && *count >= opts.MaxReplacements {
				return m
			}
			*count++
			replaced++
			return replacement
		})

		if replacedText != text {
			return strings.Replace(match, text, escapeXML(replacedText), 1)
		}
		return match
	})

	return []byte(content), replaced
}

// escapeXML escapes special XML characters
func escapeXML(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}

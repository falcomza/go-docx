package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
)

// PageNumberFormat defines the format for page numbers
type PageNumberFormat string

const (
	// PageNumDecimal uses 1, 2, 3, ... (default)
	PageNumDecimal PageNumberFormat = "decimal"
	// PageNumUpperRoman uses I, II, III, IV, ...
	PageNumUpperRoman PageNumberFormat = "upperRoman"
	// PageNumLowerRoman uses i, ii, iii, iv, ...
	PageNumLowerRoman PageNumberFormat = "lowerRoman"
	// PageNumUpperLetter uses A, B, C, ...
	PageNumUpperLetter PageNumberFormat = "upperLetter"
	// PageNumLowerLetter uses a, b, c, ...
	PageNumLowerLetter PageNumberFormat = "lowerLetter"
)

// PageNumberOptions defines options for page numbering
type PageNumberOptions struct {
	// Start is the starting page number (default: 1)
	Start int

	// Format defines the page number format (default: decimal)
	Format PageNumberFormat
}

// SetPageNumber configures page numbering for the document.
// It modifies the section properties to set the starting page number and format.
func (u *Updater) SetPageNumber(opts PageNumberOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	if opts.Start < 0 {
		return fmt.Errorf("page number start must be >= 0")
	}
	if opts.Format == "" {
		opts.Format = PageNumDecimal
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	updated, err := setPageNumberInSectPr(raw, opts)
	if err != nil {
		return fmt.Errorf("set page number: %w", err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// setPageNumberInSectPr updates or inserts pgNumType element in the document's sectPr.
func setPageNumberInSectPr(docXML []byte, opts PageNumberOptions) ([]byte, error) {
	// Build the pgNumType element
	var pgNumType string
	if opts.Start > 0 && opts.Format != "" {
		pgNumType = fmt.Sprintf(`<w:pgNumType w:start="%d" w:fmt="%s"/>`, opts.Start, opts.Format)
	} else if opts.Start > 0 {
		pgNumType = fmt.Sprintf(`<w:pgNumType w:start="%d"/>`, opts.Start)
	} else if opts.Format != "" {
		pgNumType = fmt.Sprintf(`<w:pgNumType w:fmt="%s"/>`, opts.Format)
	} else {
		return docXML, nil
	}

	// Find the last sectPr (document-level section properties)
	bodyEnd := bytes.Index(docXML, []byte("</w:body>"))
	if bodyEnd == -1 {
		return nil, fmt.Errorf("could not find </w:body> tag")
	}

	sectPrStart := bytes.LastIndex(docXML[:bodyEnd], []byte("<w:sectPr"))
	if sectPrStart == -1 {
		// No sectPr exists - create one
		newSectPr := fmt.Sprintf("<w:sectPr>%s</w:sectPr>", pgNumType)
		result := make([]byte, 0, len(docXML)+len(newSectPr))
		result = append(result, docXML[:bodyEnd]...)
		result = append(result, []byte(newSectPr)...)
		result = append(result, docXML[bodyEnd:]...)
		return result, nil
	}

	sectPrEnd := bytes.Index(docXML[sectPrStart:], []byte("</w:sectPr>"))
	if sectPrEnd == -1 {
		return nil, fmt.Errorf("malformed sectPr element")
	}
	sectPrEnd += sectPrStart + len("</w:sectPr>")

	sectPrContent := docXML[sectPrStart:sectPrEnd]

	// Check if pgNumType already exists
	existingPattern := regexp.MustCompile(`<w:pgNumType[^/]*/>`)
	if existingPattern.Match(sectPrContent) {
		// Replace existing pgNumType
		updatedSectPr := existingPattern.ReplaceAll(sectPrContent, []byte(pgNumType))
		result := make([]byte, 0, len(docXML))
		result = append(result, docXML[:sectPrStart]...)
		result = append(result, updatedSectPr...)
		result = append(result, docXML[sectPrEnd:]...)
		return result, nil
	}

	// Insert pgNumType before </w:sectPr>
	closeTag := []byte("</w:sectPr>")
	closeIdx := bytes.LastIndex(docXML[sectPrStart:sectPrEnd], closeTag)
	if closeIdx == -1 {
		return nil, fmt.Errorf("could not find </w:sectPr> closing tag")
	}
	insertPos := sectPrStart + closeIdx

	result := make([]byte, 0, len(docXML)+len(pgNumType))
	result = append(result, docXML[:insertPos]...)
	result = append(result, []byte(pgNumType)...)
	result = append(result, docXML[insertPos:]...)

	return result, nil
}

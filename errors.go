package docxupdater

import "fmt"

// ErrorCode represents specific error conditions
type ErrorCode string

const (
	// File-related errors
	ErrCodeInvalidFile      ErrorCode = "INVALID_FILE"
	ErrCodeFileNotFound     ErrorCode = "FILE_NOT_FOUND"
	ErrCodeFileCorrupted    ErrorCode = "FILE_CORRUPTED"
	ErrCodeFileTooLarge     ErrorCode = "FILE_TOO_LARGE"
	ErrCodePermissionDenied ErrorCode = "PERMISSION_DENIED"

	// Chart-related errors
	ErrCodeChartNotFound    ErrorCode = "CHART_NOT_FOUND"
	ErrCodeInvalidChartData ErrorCode = "INVALID_CHART_DATA"
	ErrCodeChartCreation    ErrorCode = "CHART_CREATION"

	// Table-related errors
	ErrCodeInvalidTableData ErrorCode = "INVALID_TABLE_DATA"
	ErrCodeTableCreation    ErrorCode = "TABLE_CREATION"

	// Image-related errors
	ErrCodeImageNotFound ErrorCode = "IMAGE_NOT_FOUND"
	ErrCodeImageFormat   ErrorCode = "IMAGE_FORMAT"
	ErrCodeImageTooLarge ErrorCode = "IMAGE_TOO_LARGE"

	// Text-related errors
	ErrCodeTextNotFound   ErrorCode = "TEXT_NOT_FOUND"
	ErrCodeInvalidRegex   ErrorCode = "INVALID_REGEX"
	ErrCodeReplaceFailure ErrorCode = "REPLACE_FAILURE"

	// XML-related errors
	ErrCodeXMLParse   ErrorCode = "XML_PARSE"
	ErrCodeXMLWrite   ErrorCode = "XML_WRITE"
	ErrCodeInvalidXML ErrorCode = "INVALID_XML"

	// Relationship errors
	ErrCodeRelationship ErrorCode = "RELATIONSHIP"
	ErrCodeRelNotFound  ErrorCode = "RELATIONSHIP_NOT_FOUND"

	// Content type errors
	ErrCodeContentType ErrorCode = "CONTENT_TYPE"

	// Validation errors
	ErrCodeValidation      ErrorCode = "VALIDATION"
	ErrCodeMissingRequired ErrorCode = "MISSING_REQUIRED"
	ErrCodeInvalidValue    ErrorCode = "INVALID_VALUE"

	// Document structure errors
	ErrCodeNoDocument       ErrorCode = "NO_DOCUMENT"
	ErrCodeNoBody           ErrorCode = "NO_BODY"
	ErrCodeInvalidStructure ErrorCode = "INVALID_STRUCTURE"

	// Hyperlink errors
	ErrCodeHyperlinkCreation ErrorCode = "HYPERLINK_CREATION"
	ErrCodeInvalidURL        ErrorCode = "INVALID_URL"

	// Header/Footer errors
	ErrCodeHeaderFooter ErrorCode = "HEADER_FOOTER"
)

// DocxError provides structured error information
type DocxError struct {
	Code    ErrorCode
	Message string
	Err     error
	Context map[string]any
}

// Error implements the error interface
func (e *DocxError) Error() string {
	if e.Err != nil {
		return fmt.Sprintf("%s: %s: %v", e.Code, e.Message, e.Err)
	}
	return fmt.Sprintf("%s: %s", e.Code, e.Message)
}

// Unwrap returns the wrapped error
func (e *DocxError) Unwrap() error {
	return e.Err
}

// WithContext adds context to the error
func (e *DocxError) WithContext(key string, value any) *DocxError {
	if e.Context == nil {
		e.Context = make(map[string]any)
	}
	e.Context[key] = value
	return e
}

// Constructor helpers for common errors

// NewChartNotFoundError creates an error for when a chart is not found
func NewChartNotFoundError(index int) error {
	return &DocxError{
		Code:    ErrCodeChartNotFound,
		Message: "chart not found",
		Context: map[string]any{"index": index},
	}
}

// NewInvalidChartDataError creates an error for invalid chart data
func NewInvalidChartDataError(reason string) error {
	return &DocxError{
		Code:    ErrCodeInvalidChartData,
		Message: reason,
	}
}

// NewImageNotFoundError creates an error for when an image file is not found
func NewImageNotFoundError(path string) error {
	return &DocxError{
		Code:    ErrCodeImageNotFound,
		Message: "image file not found",
		Context: map[string]any{"path": path},
	}
}

// NewImageFormatError creates an error for unsupported image formats
func NewImageFormatError(format string) error {
	return &DocxError{
		Code:    ErrCodeImageFormat,
		Message: "unsupported image format",
		Context: map[string]any{"format": format},
	}
}

// NewTextNotFoundError creates an error for when text is not found in document
func NewTextNotFoundError(text string) error {
	return &DocxError{
		Code:    ErrCodeTextNotFound,
		Message: "text not found in document",
		Context: map[string]any{"text": text},
	}
}

// NewInvalidRegexError creates an error for invalid regex patterns
func NewInvalidRegexError(pattern string, err error) error {
	return &DocxError{
		Code:    ErrCodeInvalidRegex,
		Message: "invalid regular expression pattern",
		Err:     err,
		Context: map[string]any{"pattern": pattern},
	}
}

// NewXMLParseError creates an error for XML parsing failures
func NewXMLParseError(file string, err error) error {
	return &DocxError{
		Code:    ErrCodeXMLParse,
		Message: "failed to parse XML",
		Err:     err,
		Context: map[string]any{"file": file},
	}
}

// NewXMLWriteError creates an error for XML writing failures
func NewXMLWriteError(file string, err error) error {
	return &DocxError{
		Code:    ErrCodeXMLWrite,
		Message: "failed to write XML",
		Err:     err,
		Context: map[string]any{"file": file},
	}
}

// NewRelationshipError creates an error for relationship issues
func NewRelationshipError(reason string, err error) error {
	return &DocxError{
		Code:    ErrCodeRelationship,
		Message: reason,
		Err:     err,
	}
}

// NewValidationError creates an error for validation failures
func NewValidationError(field, reason string) error {
	return &DocxError{
		Code:    ErrCodeValidation,
		Message: reason,
		Context: map[string]any{"field": field},
	}
}

// NewFileNotFoundError creates an error for missing files
func NewFileNotFoundError(path string) error {
	return &DocxError{
		Code:    ErrCodeFileNotFound,
		Message: "file not found",
		Context: map[string]any{"path": path},
	}
}

// NewInvalidFileError creates an error for invalid DOCX files
func NewInvalidFileError(reason string, err error) error {
	return &DocxError{
		Code:    ErrCodeInvalidFile,
		Message: reason,
		Err:     err,
	}
}

// NewHyperlinkError creates an error for hyperlink creation failures
func NewHyperlinkError(reason string, err error) error {
	return &DocxError{
		Code:    ErrCodeHyperlinkCreation,
		Message: reason,
		Err:     err,
	}
}

// NewInvalidURLError creates an error for invalid URLs
func NewInvalidURLError(url string) error {
	return &DocxError{
		Code:    ErrCodeInvalidURL,
		Message: "invalid URL format",
		Context: map[string]any{"url": url},
	}
}

// NewHeaderFooterError creates an error for header/footer operations
func NewHeaderFooterError(reason string, err error) error {
	return &DocxError{
		Code:    ErrCodeHeaderFooter,
		Message: reason,
		Err:     err,
	}
}

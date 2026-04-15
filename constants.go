package godocx

import "regexp"

// FontSizeHalfPointsFactor is the multiplier to convert a font size expressed in
// typographic points to the half-point units required by the Open XML w:sz and
// w:szCs attributes (ECMA-376 Part 1 §17.3.2.38).
const FontSizeHalfPointsFactor = 2

// OpenXML constants for chart drawings
const (
	// ChartAnchorIDBase is the base value for anchor IDs in chart drawings
	ChartAnchorIDBase = 0x30000000

	// ChartEditIDBase is the base value for edit IDs in chart drawings
	ChartEditIDBase = 0x0D000000

	// ChartIDIncrement is the increment per chart to ensure ID uniqueness
	ChartIDIncrement = 0x1000
)

// OpenXML constants for image drawings
const (
	// ImageAnchorIDBase is the base value for anchor IDs in image drawings
	ImageAnchorIDBase = 0x50000000

	// ImageEditIDBase is the base value for edit IDs in image drawings
	ImageEditIDBase = 0x0E000000

	// ImageIDIncrement is the increment per image to ensure ID uniqueness
	ImageIDIncrement = 0x1000

	// EMUsPerInch is the number of English Metric Units (EMUs) per inch
	// 1 inch = 914400 EMUs (used for image sizing in OpenXML)
	EMUsPerInch = 914400

	// DefaultImageDPI is the default DPI for image dimensions
	DefaultImageDPI = 96
)

// Package-level compiled regular expressions for performance
var (
	// chartFilePattern matches chart XML filenames (e.g., chart1.xml, chart2.xml)
	chartFilePattern = regexp.MustCompile(`^chart(\d+)\.xml$`)

	// imageFilePattern matches image filenames in media folder (e.g., image1.png, image2.jpg)
	imageFilePattern = regexp.MustCompile(`^image(\d+)\.\w+$`)

	// docPrIDPattern matches docPr id attributes in document.xml
	docPrIDPattern = regexp.MustCompile(`docPr id="(\d+)"`)

	// bookmarkIDPattern matches w:id attributes on bookmark start elements only.
	// Using a narrowly-scoped pattern prevents inflating the ID counter with
	// w:id attributes on unrelated elements (w:tc, w:comment, w:ins, etc.).
	bookmarkIDPattern = regexp.MustCompile(`<w:bookmarkStart[^>]+w:id="(\d+)"`)

	// relIDPattern matches relationship IDs (e.g., rId1, rId2)
	relIDPattern = regexp.MustCompile(`^rId(\d+)$`)

	// textRunPattern matches Word text runs (<w:t ...>...</w:t>)
	textRunPattern = regexp.MustCompile(`<w:t(?:\s[^>]*)?(>.*?</w:t>)`)

	// textContentPattern extracts text from a Word text run
	textContentPattern = regexp.MustCompile(`<w:t(?:\s[^>]*)?>(.*)</w:t>`)

	// extractTextPattern matches visible text inside <w:t> elements.
	// Uses [ \t] (not \s) to avoid matching newlines within the tag.
	extractTextPattern = regexp.MustCompile(`<w:t(?:[ \t][^>]*)?>([^<]*)</w:t>`)

	// extractParaPattern matches full <w:p> paragraph elements.
	extractParaPattern = regexp.MustCompile(`(?s)<w:p[^>]*>.*?</w:p>`)

	// extractTablePattern matches full <w:tbl> table elements.
	extractTablePattern = regexp.MustCompile(`(?s)<w:tbl>.*?</w:tbl>`)

	// extractRowPattern matches full <w:tr> table-row elements.
	extractRowPattern = regexp.MustCompile(`(?s)<w:tr[^>]*>.*?</w:tr>`)

	// extractCellPattern matches full <w:tc> table-cell elements.
	extractCellPattern = regexp.MustCompile(`(?s)<w:tc>.*?</w:tc>`)

	// runBlockPattern matches a complete <w:r> run element (with optional attributes).
	// w:r elements are leaf nodes inside a paragraph — they never nest — so regex
	// is safe. Used for run-normalization before placeholder substitution.
	runBlockPattern = regexp.MustCompile(`(?s)<w:r(?:\s[^>]*)?>.*?</w:r>`)

	// runRprPattern extracts the <w:rPr> run-properties block from within a run.
	// Handles both non-empty (<w:rPr>…</w:rPr>) and self-closing (<w:rPr/>) forms.
	runRprPattern = regexp.MustCompile(`(?s)<w:rPr(?:\s[^>]*)?(?:/>|>.*?</w:rPr>)`)
)

// OpenXML namespace URIs
const (
	RelationshipsNS  = "http://schemas.openxmlformats.org/package/2006/relationships"
	OfficeDocumentNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	DrawingMLNS      = "http://schemas.openxmlformats.org/drawingml/2006/main"
	ChartNS          = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	SpreadsheetMLNS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)

// VML and Office namespace URIs (used for OLE embedded objects)
const (
	VMLNamespace    = "urn:schemas-microsoft-com:vml"
	OfficeNamespace = "urn:schemas-microsoft-com:office:office"
)

// OLE/embedded-object constants
const (
	// OLEPackageRelType is the relationship type for an embedded OpenXML package
	// (e.g., an Excel workbook embedded as a clickable object in a Word document).
	OLEPackageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"

	// XLSXContentType is the MIME type for .xlsx files.
	XLSXContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

	// OLEProgIDExcel is the OLE program identifier for Excel 2007+ workbooks.
	OLEProgIDExcel = "Excel.Sheet.12"

	// DefaultEmbedWidthPt is the default display width of an embedded object in points.
	DefaultEmbedWidthPt = 95

	// DefaultEmbedHeightPt is the default display height of an embedded object in points.
	DefaultEmbedHeightPt = 75
)

// embeddingFilePattern matches numbered embedding filenames (e.g., embedding1.xlsx).
var embeddingFilePattern = regexp.MustCompile(`^embedding(\d+)\.xlsx$`)


// OpenXML content types
const (
	ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

	// DocxMainContentType is the document body content type for .docx files.
	DocxMainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

	// DotxMainContentType is the document body content type for .dotx template files.
	// New() automatically promotes this to DocxMainContentType so templates can be
	// used as input without any special handling by the caller.
	DotxMainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
	ImageJPEGType    = "image/jpeg"
	ImagePNGType     = "image/png"
	ImageGIFType     = "image/gif"
	ImageBMPType     = "image/bmp"
	ImageTIFFType    = "image/tiff"
)

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

// defaultExcelIconPNG is a minimal 95×75 green-grid PNG icon used when no custom icon is provided.
// 376 bytes
var defaultExcelIconPNG = []byte{
	0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
	0x00, 0x00, 0x00, 0x5f, 0x00, 0x00, 0x00, 0x4b, 0x08, 0x02, 0x00, 0x00, 0x00, 0x99, 0x69, 0xcc,
	0x2f, 0x00, 0x00, 0x01, 0x3f, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0xec, 0xd8, 0x31, 0x52, 0x84,
	0x30, 0x18, 0x40, 0xe1, 0xe8, 0x6c, 0x6b, 0x67, 0x63, 0x29, 0x87, 0x91, 0x03, 0xc9, 0x31, 0xb8,
	0x10, 0x97, 0xa1, 0xa5, 0xa1, 0xe3, 0x02, 0x16, 0xce, 0x38, 0xe1, 0x27, 0x79, 0x8a, 0x16, 0x8a,
	0x79, 0x5f, 0xe7, 0x90, 0x71, 0x77, 0xde, 0x6c, 0x42, 0xe6, 0xbf, 0x3d, 0xbf, 0xbe, 0x24, 0x55,
	0xdc, 0xff, 0xf6, 0x17, 0xf8, 0xd3, 0xac, 0x43, 0xac, 0x43, 0xac, 0x43, 0x62, 0x9d, 0x79, 0x9c,
	0x6a, 0x4b, 0xe7, 0x71, 0xda, 0x96, 0xf5, 0xeb, 0xff, 0xfa, 0x8a, 0x8b, 0x83, 0x58, 0xa7, 0x1b,
	0xfa, 0x62, 0xa0, 0x79, 0x9c, 0xba, 0xa1, 0xff, 0xf6, 0xc7, 0x5c, 0x54, 0x61, 0x67, 0x1d, 0x03,
	0xb5, 0x99, 0xa6, 0x7a, 0xee, 0xe4, 0x81, 0x9a, 0x4d, 0x93, 0x52, 0xba, 0xd5, 0x1e, 0x7c, 0x04,
	0x6a, 0x36, 0x8d, 0xef, 0xac, 0x4f, 0x54, 0xeb, 0xbc, 0x6f, 0xa8, 0xda, 0x21, 0xdd, 0x88, 0x72,
	0x9d, 0xfc, 0xac, 0x69, 0x39, 0x50, 0xa1, 0xce, 0xf1, 0x18, 0x6e, 0x36, 0x50, 0xe1, 0x36, 0x58,
	0x3c, 0x86, 0xdb, 0x0c, 0x74, 0x17, 0x26, 0x18, 0x3f, 0xb9, 0x59, 0xfe, 0x03, 0x0f, 0x4f, 0x8f,
	0xf9, 0x9f, 0xf1, 0x8d, 0x1e, 0x1e, 0x07, 0xdb, 0xb2, 0xf2, 0x82, 0xab, 0x2f, 0x0e, 0x7c, 0xa3,
	0x13, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb,
	0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0x27, 0x18, 0x3b, 0x4e, 0x30, 0x4e, 0x70,
	0x67, 0x11, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10,
	0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0x27, 0x18, 0x3b, 0x4e, 0x30, 0x4e,
	0x70, 0x67, 0x11, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb,
	0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0x27, 0x18, 0x3b, 0x4e, 0x30,
	0x4e, 0x70, 0x67, 0x11, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10, 0xeb, 0x10,
	0xeb, 0x90, 0x78, 0x57, 0x56, 0xce, 0xdf, 0x0e, 0xb1, 0x0e, 0xb1, 0x0e, 0xb1, 0x0e, 0x79, 0x0b,
	0x00, 0x00, 0xff, 0xff, 0x5e, 0xa4, 0x8e, 0xa3, 0x7f, 0xa7, 0xd0, 0x44, 0x00, 0x00, 0x00, 0x00,
	0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
}

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

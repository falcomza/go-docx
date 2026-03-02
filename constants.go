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

	// bookmarkIDPattern matches bookmark id attributes (w:bookmarkStart and w:bookmarkEnd)
	bookmarkIDPattern = regexp.MustCompile(`w:id="(\d+)"`)

	// relIDPattern matches relationship IDs (e.g., rId1, rId2)
	relIDPattern = regexp.MustCompile(`^rId(\d+)$`)

	// textRunPattern matches Word text runs (<w:t ...>...</w:t>)
	textRunPattern = regexp.MustCompile(`<w:t(?:\s[^>]*)?(>.*?</w:t>)`)

	// textContentPattern extracts text from a Word text run
	textContentPattern = regexp.MustCompile(`<w:t(?:\s[^>]*)?>(.*)</w:t>`)
)

// OpenXML namespace URIs
const (
	RelationshipsNS  = "http://schemas.openxmlformats.org/package/2006/relationships"
	OfficeDocumentNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	DrawingMLNS      = "http://schemas.openxmlformats.org/drawingml/2006/main"
	ChartNS          = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	SpreadsheetMLNS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)

// OpenXML content types
const (
	ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	ImageJPEGType    = "image/jpeg"
	ImagePNGType     = "image/png"
	ImageGIFType     = "image/gif"
	ImageBMPType     = "image/bmp"
	ImageTIFFType    = "image/tiff"
)

package docxupdater

// ChartData defines chart categories and series values.
type ChartData struct {
	Categories []string
	Series     []SeriesData
	// Optional titles
	ChartTitle        string // Main chart title
	CategoryAxisTitle string // X-axis title
	ValueAxisTitle    string // Y-axis title
}

// SeriesData defines one chart series.
type SeriesData struct {
	Name   string
	Values []float64
	Color  string // Hex color code (e.g., "FF0000" for red) - optional
}

// ImageOptions defines options for image insertion
type ImageOptions struct {
	// Path to the image file (required)
	Path string

	// Width in pixels (optional - if only width is set, height is calculated proportionally)
	Width int

	// Height in pixels (optional - if only height is set, width is calculated proportionally)
	Height int

	// Alternative text for accessibility (optional)
	AltText string

	// Position where to insert the image
	Position InsertPosition

	// Anchor text for position-based insertion (for PositionAfterText/PositionBeforeText)
	Anchor string

	// Caption options (nil for no caption)
	Caption *CaptionOptions
}

// ImageDimensions stores image width and height in pixels
type ImageDimensions struct {
	Width  int
	Height int
}

// SectionBreakType defines the type of section break
type SectionBreakType string

const (
	// SectionBreakNextPage starts the new section on the next page
	SectionBreakNextPage SectionBreakType = "nextPage"
	// SectionBreakContinuous starts the new section on the same page
	SectionBreakContinuous SectionBreakType = "continuous"
	// SectionBreakEvenPage starts the new section on the next even page
	SectionBreakEvenPage SectionBreakType = "evenPage"
	// SectionBreakOddPage starts the new section on the next odd page
	SectionBreakOddPage SectionBreakType = "oddPage"
)

// BreakOptions defines options for inserting breaks
type BreakOptions struct {
	// Position where to insert the break
	Position InsertPosition

	// Anchor text for position-based insertion (for PositionAfterText/PositionBeforeText)
	Anchor string

	// Type of section break (only used for section breaks)
	SectionType SectionBreakType

	// Page layout settings for the section (optional)
	PageLayout *PageLayoutOptions
}

// PageOrientation defines page orientation
type PageOrientation string

const (
	// OrientationPortrait sets page to portrait mode (taller than wide)
	OrientationPortrait PageOrientation = "portrait"
	// OrientationLandscape sets page to landscape mode (wider than tall)
	OrientationLandscape PageOrientation = "landscape"
)

// PageLayoutOptions defines page layout settings for a section
type PageLayoutOptions struct {
	// Page dimensions in twips (1/1440 inch)
	// Use predefined page sizes or custom values
	PageWidth  int // Width in twips (e.g., 12240 for 8.5")
	PageHeight int // Height in twips (e.g., 15840 for 11")

	// Page orientation
	Orientation PageOrientation

	// Margins in twips (1/1440 inch)
	MarginTop    int // Top margin
	MarginRight  int // Right margin
	MarginBottom int // Bottom margin
	MarginLeft   int // Left margin
	MarginHeader int // Header distance from edge
	MarginFooter int // Footer distance from edge
	MarginGutter int // Gutter margin for binding
}

// Page size constants in twips (1/1440 inch)
const (
	// US Letter: 8.5" × 11"
	PageWidthLetter  = 12240
	PageHeightLetter = 15840

	// US Legal: 8.5" × 14"
	PageWidthLegal  = 12240
	PageHeightLegal = 20160

	// A4: 210mm × 297mm (8.27" × 11.69")
	PageWidthA4  = 11906
	PageHeightA4 = 16838

	// A3: 297mm × 420mm (11.69" × 16.54")
	PageWidthA3  = 16838
	PageHeightA3 = 23811

	// Tabloid/Ledger: 11" × 17"
	PageWidthTabloid  = 15840
	PageHeightTabloid = 24480

	// Default margin: 1 inch
	MarginDefault = 1440

	// Narrow margin: 0.5 inch
	MarginNarrow = 720

	// Wide margin: 1.5 inch
	MarginWide = 2160

	// Default header/footer margin: 0.5 inch
	MarginHeaderFooterDefault = 720
)

// Helper functions to create common page layouts

// PageLayoutLetterPortrait creates a US Letter portrait layout with 1" margins
func PageLayoutLetterPortrait() *PageLayoutOptions {
	return &PageLayoutOptions{
		PageWidth:    PageWidthLetter,
		PageHeight:   PageHeightLetter,
		Orientation:  OrientationPortrait,
		MarginTop:    MarginDefault,
		MarginRight:  MarginDefault,
		MarginBottom: MarginDefault,
		MarginLeft:   MarginDefault,
		MarginHeader: MarginHeaderFooterDefault,
		MarginFooter: MarginHeaderFooterDefault,
		MarginGutter: 0,
	}
}

// PageLayoutLetterLandscape creates a US Letter landscape layout with 1" margins
func PageLayoutLetterLandscape() *PageLayoutOptions {
	return &PageLayoutOptions{
		PageWidth:    PageHeightLetter, // Swapped for landscape
		PageHeight:   PageWidthLetter,  // Swapped for landscape
		Orientation:  OrientationLandscape,
		MarginTop:    MarginDefault,
		MarginRight:  MarginDefault,
		MarginBottom: MarginDefault,
		MarginLeft:   MarginDefault,
		MarginHeader: MarginHeaderFooterDefault,
		MarginFooter: MarginHeaderFooterDefault,
		MarginGutter: 0,
	}
}

// PageLayoutA4Portrait creates an A4 portrait layout with default margins
func PageLayoutA4Portrait() *PageLayoutOptions {
	return &PageLayoutOptions{
		PageWidth:    PageWidthA4,
		PageHeight:   PageHeightA4,
		Orientation:  OrientationPortrait,
		MarginTop:    MarginDefault,
		MarginRight:  MarginDefault,
		MarginBottom: MarginDefault,
		MarginLeft:   MarginDefault,
		MarginHeader: MarginHeaderFooterDefault,
		MarginFooter: MarginHeaderFooterDefault,
		MarginGutter: 0,
	}
}

// PageLayoutA4Landscape creates an A4 landscape layout with default margins
func PageLayoutA4Landscape() *PageLayoutOptions {
	return &PageLayoutOptions{
		PageWidth:    PageHeightA4, // Swapped for landscape
		PageHeight:   PageWidthA4,  // Swapped for landscape
		Orientation:  OrientationLandscape,
		MarginTop:    MarginDefault,
		MarginRight:  MarginDefault,
		MarginBottom: MarginDefault,
		MarginLeft:   MarginDefault,
		MarginHeader: MarginHeaderFooterDefault,
		MarginFooter: MarginHeaderFooterDefault,
		MarginGutter: 0,
	}
}

// PageLayoutLegalPortrait creates a US Legal portrait layout with 1" margins
func PageLayoutLegalPortrait() *PageLayoutOptions {
	return &PageLayoutOptions{
		PageWidth:    PageWidthLegal,
		PageHeight:   PageHeightLegal,
		Orientation:  OrientationPortrait,
		MarginTop:    MarginDefault,
		MarginRight:  MarginDefault,
		MarginBottom: MarginDefault,
		MarginLeft:   MarginDefault,
		MarginHeader: MarginHeaderFooterDefault,
		MarginFooter: MarginHeaderFooterDefault,
		MarginGutter: 0,
	}
}

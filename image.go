package godocx

import (
	"bytes"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// InsertImage inserts an image into the document with optional proportional sizing
func (u *Updater) InsertImage(opts ImageOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Path == "" {
		return fmt.Errorf("image path cannot be empty")
	}

	// Check if the image file exists
	if _, err := os.Stat(opts.Path); os.IsNotExist(err) {
		return fmt.Errorf("image file not found: %s", opts.Path)
	}

	// Get actual image dimensions from file
	actualDims, err := getImageDimensions(opts.Path)
	if err != nil {
		return fmt.Errorf("get image dimensions: %w", err)
	}

	// Calculate final dimensions (with proportions if needed)
	finalDims := calculateProportionalDimensions(actualDims, opts.Width, opts.Height)

	// Get next image index
	imageIndex, err := u.getNextImageIndex()
	if err != nil {
		return fmt.Errorf("get next image index: %w", err)
	}

	// Determine content type and file extension
	contentType := getImageContentType(opts.Path)
	ext := strings.ToLower(filepath.Ext(opts.Path))

	// Copy image to media folder
	imageFileName := fmt.Sprintf("image%d%s", imageIndex, ext)
	if err := u.copyImageToMedia(opts.Path, imageFileName); err != nil {
		return fmt.Errorf("copy image to media: %w", err)
	}

	// Add relationship for the image
	relId, err := u.addImageRelationship(imageFileName)
	if err != nil {
		return fmt.Errorf("add image relationship: %w", err)
	}

	// Add content type for the image
	if err := u.addImageContentType(ext, contentType); err != nil {
		return fmt.Errorf("add image content type: %w", err)
	}

	// Generate image drawing XML
	imageXML, err := u.generateImageDrawingXML(imageIndex, relId, finalDims, opts.AltText)
	if err != nil {
		return fmt.Errorf("generate image drawing: %w", err)
	}

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Handle caption if provided
	var contentToInsert []byte
	if opts.Caption != nil {
		// Validate caption options
		if err := ValidateCaptionOptions(opts.Caption); err != nil {
			return fmt.Errorf("invalid caption options: %w", err)
		}

		// Default to Figure caption type for images
		if opts.Caption.Type == "" {
			opts.Caption.Type = CaptionFigure
		}

		// Generate caption XML
		captionXML := generateCaptionXML(*opts.Caption)

		// Combine image and caption based on position
		contentToInsert = insertCaptionWithElement(raw, captionXML, imageXML, opts.Caption.Position)
	} else {
		contentToInsert = imageXML
	}

	// Insert image at the specified position
	updated, err := insertImageAtPosition(raw, contentToInsert, opts)
	if err != nil {
		return fmt.Errorf("insert image: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// getImageDimensions reads the image file and returns its dimensions in pixels
func getImageDimensions(path string) (ImageDimensions, error) {
	file, err := os.Open(path)
	if err != nil {
		return ImageDimensions{}, fmt.Errorf("open image: %w", err)
	}
	defer file.Close()

	config, _, err := image.DecodeConfig(file)
	if err != nil {
		return ImageDimensions{}, fmt.Errorf("decode image config: %w", err)
	}

	return ImageDimensions{
		Width:  config.Width,
		Height: config.Height,
	}, nil
}

// calculateProportionalDimensions calculates final dimensions maintaining aspect ratio
// If both width and height are provided, uses them as-is
// If only width is provided, calculates height proportionally
// If only height is provided, calculates width proportionally
// If neither is provided, uses actual image dimensions
func calculateProportionalDimensions(actual ImageDimensions, requestedWidth, requestedHeight int) ImageDimensions {
	// If both dimensions are specified, use them
	if requestedWidth > 0 && requestedHeight > 0 {
		return ImageDimensions{
			Width:  requestedWidth,
			Height: requestedHeight,
		}
	}

	// If only width is specified, calculate height proportionally
	if requestedWidth > 0 {
		ratio := float64(actual.Height) / float64(actual.Width)
		return ImageDimensions{
			Width:  requestedWidth,
			Height: int(float64(requestedWidth) * ratio),
		}
	}

	// If only height is specified, calculate width proportionally
	if requestedHeight > 0 {
		ratio := float64(actual.Width) / float64(actual.Height)
		return ImageDimensions{
			Width:  int(float64(requestedHeight) * ratio),
			Height: requestedHeight,
		}
	}

	// Neither specified, use actual dimensions
	return actual
}

// convertPixelsToEMUs converts pixels to English Metric Units (EMUs)
// EMUs are used by OpenXML for positioning and sizing
func convertPixelsToEMUs(pixels int) int64 {
	// Convert pixels to inches, then to EMUs
	// Default to 96 DPI
	return int64(float64(pixels) * float64(EMUsPerInch) / float64(DefaultImageDPI))
}

// getImageContentType returns the MIME type for the image based on file extension
func getImageContentType(path string) string {
	ext := strings.ToLower(filepath.Ext(path))
	switch ext {
	case ".jpg", ".jpeg":
		return ImageJPEGType
	case ".png":
		return ImagePNGType
	case ".gif":
		return ImageGIFType
	case ".bmp":
		return ImageBMPType
	case ".tif", ".tiff":
		return ImageTIFFType
	default:
		return ImagePNGType // default to PNG
	}
}

// getNextImageIndex finds the next available image index by scanning the media folder
func (u *Updater) getNextImageIndex() (int, error) {
	mediaPath := filepath.Join(u.tempDir, "word", "media")

	// Create media folder if it doesn't exist
	if err := os.MkdirAll(mediaPath, 0o755); err != nil {
		return 0, fmt.Errorf("create media folder: %w", err)
	}

	entries, err := os.ReadDir(mediaPath)
	if err != nil {
		return 0, fmt.Errorf("read media folder: %w", err)
	}

	maxIndex := 0
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		matches := imageFilePattern.FindStringSubmatch(entry.Name())
		if len(matches) > 1 {
			index, err := strconv.Atoi(matches[1])
			if err == nil && index > maxIndex {
				maxIndex = index
			}
		}
	}

	return maxIndex + 1, nil
}

// generateImageDrawingXML creates the inline drawing XML for an image
func (u *Updater) generateImageDrawingXML(imageIndex int, relId string, dims ImageDimensions, altText string) ([]byte, error) {
	// Get a unique docPr ID
	docPrId, err := u.getNextDocPrId()
	if err != nil {
		return nil, fmt.Errorf("get next docPr id: %w", err)
	}

	// Convert dimensions to EMUs
	widthEMU := convertPixelsToEMUs(dims.Width)
	heightEMU := convertPixelsToEMUs(dims.Height)

	// Generate unique IDs
	anchorId := ImageAnchorIDBase + uint32(imageIndex)*ImageIDIncrement
	editId := ImageEditIDBase + uint32(imageIndex)*ImageIDIncrement

	// Default alt text if not provided
	if altText == "" {
		altText = fmt.Sprintf("Picture %d", imageIndex)
	}

	template := `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="%08X" wp14:editId="%08X"><wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="%d" name="Picture %d" descr="%s"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="%d" name="Picture %d" descr="%s"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

	return fmt.Appendf(nil, template,
		anchorId, editId, widthEMU, heightEMU,
		docPrId, imageIndex, xmlEscape(altText),
		docPrId, imageIndex, xmlEscape(altText),
		relId,
		widthEMU, heightEMU), nil
}

// addImageRelationship adds a relationship for the image to document.xml.rels
func (u *Updater) addImageRelationship(imageFileName string) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document relationships: %w", err)
	}

	// Get next relationship ID
	nextRelId, err := u.getNextDocumentRelId()
	if err != nil {
		return "", err
	}

	// Add relationship for the image
	insert := fmt.Sprintf("\n  <Relationship Id=\"%s\" Type=\"%s/image\" Target=\"media/%s\"/>\n",
		nextRelId, OfficeDocumentNS, imageFileName)

	closer := []byte("</Relationships>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return "", fmt.Errorf("invalid document.xml.rels: missing </Relationships>")
	}

	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])

	if err := os.WriteFile(relsPath, result, 0o644); err != nil {
		return "", fmt.Errorf("write relationships: %w", err)
	}

	return nextRelId, nil
}

// addImageContentType adds or ensures the image extension is registered in [Content_Types].xml
func (u *Updater) addImageContentType(ext, contentType string) error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	raw, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	// Remove leading dot from extension
	ext = strings.TrimPrefix(ext, ".")

	// Check if Default element for this extension already exists
	checkPattern := fmt.Sprintf(`Extension="%s"`, ext)
	if bytes.Contains(raw, []byte(checkPattern)) {
		return nil // already present
	}

	// Add Default element for the image type
	insert := fmt.Sprintf("\n  <Default Extension=\"%s\" ContentType=\"%s\"/>\n", ext, contentType)

	// Insert before </Types>
	closer := []byte("</Types>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return fmt.Errorf("invalid [Content_Types].xml: missing </Types>")
	}

	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])

	return os.WriteFile(contentTypesPath, result, 0o644)
}

// copyImageToMedia copies the image file to the word/media folder
func (u *Updater) copyImageToMedia(srcPath, destFileName string) error {
	mediaPath := filepath.Join(u.tempDir, "word", "media")

	// Ensure media directory exists
	if err := os.MkdirAll(mediaPath, 0o755); err != nil {
		return fmt.Errorf("create media directory: %w", err)
	}

	// Open source file
	srcFile, err := os.Open(srcPath)
	if err != nil {
		return fmt.Errorf("open source file: %w", err)
	}
	defer srcFile.Close()

	// Create destination file
	destPath := filepath.Join(mediaPath, destFileName)
	destFile, err := os.Create(destPath)
	if err != nil {
		return fmt.Errorf("create destination file: %w", err)
	}
	defer destFile.Close()

	// Copy file content
	if _, err := io.Copy(destFile, srcFile); err != nil {
		return fmt.Errorf("copy file: %w", err)
	}

	return nil
}

// insertImageAtPosition inserts the image XML at the specified position in document.xml
func insertImageAtPosition(raw []byte, imageXML []byte, opts ImageOptions) ([]byte, error) {

	switch opts.Position {
	case PositionBeginning:
		return insertAtBodyStart(raw, imageXML)

	case PositionEnd:
		return insertAtBodyEnd(raw, imageXML)

	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionAfterText")
		}
		return insertAfterText(raw, imageXML, opts.Anchor)

	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, fmt.Errorf("anchor text required for PositionBeforeText")
		}
		return insertBeforeText(raw, imageXML, opts.Anchor)

	default:
		return nil, fmt.Errorf("invalid position: %d", opts.Position)
	}
}

// insertAfterAnchor and insertBeforeAnchor removed in favor of paragraph-aware helpers.

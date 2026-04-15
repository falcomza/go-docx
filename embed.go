package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
)

// InsertEmbeddedObject inserts an OLE embedded object (e.g., an Excel workbook)
// into the document. The object is displayed as a clickable icon; double-clicking
// it in Word/LibreOffice opens the file in the associated application.
func (u *Updater) InsertEmbeddedObject(opts EmbeddedObjectOptions) error {
	opts = applyEmbedDefaults(opts)

	// Validate anchor before any I/O so callers get fast feedback on bad input.
	if (opts.Position == PositionAfterText || opts.Position == PositionBeforeText) && opts.Anchor == "" {
		return fmt.Errorf("anchor text required for position %d", opts.Position)
	}

	fileBytes, err := resolveEmbedFileBytes(opts)
	if err != nil {
		return fmt.Errorf("resolve embed file: %w", err)
	}

	iconBytes := resolveIconBytes(opts)

	embIdx := u.findNextEmbeddingIndex()

	// getNextImageIndex also creates word/media/ if it doesn't exist.
	imgIdx, err := u.getNextImageIndex()
	if err != nil {
		return fmt.Errorf("get next image index: %w", err)
	}

	docPrId, err := u.getNextDocPrId()
	if err != nil {
		return fmt.Errorf("get next docPr id: %w", err)
	}

	embDir := filepath.Join(u.tempDir, "word", "embeddings")
	if err := os.MkdirAll(embDir, 0o755); err != nil {
		return fmt.Errorf("create embeddings dir: %w", err)
	}
	xlsxFileName := fmt.Sprintf("embedding%d.xlsx", embIdx)
	if err := atomicWriteFile(filepath.Join(embDir, xlsxFileName), fileBytes, 0o644); err != nil {
		return fmt.Errorf("write embedded xlsx: %w", err)
	}

	iconFileName := fmt.Sprintf("image%d.png", imgIdx)
	if err := atomicWriteFile(filepath.Join(u.tempDir, "word", "media", iconFileName), iconBytes, 0o644); err != nil {
		return fmt.Errorf("write icon image: %w", err)
	}

	xlsxRelID, err := u.addDocumentRelationship(OLEPackageRelType, "embeddings/"+xlsxFileName)
	if err != nil {
		return fmt.Errorf("add xlsx relationship: %w", err)
	}

	imageRelID, err := u.addImageRelationship(iconFileName)
	if err != nil {
		return fmt.Errorf("add image relationship: %w", err)
	}

	if err := u.addImageContentType(".xlsx", XLSXContentType); err != nil {
		return fmt.Errorf("add xlsx content type: %w", err)
	}
	if err := u.addImageContentType(".png", ImagePNGType); err != nil {
		return fmt.Errorf("add icon content type: %w", err)
	}

	shapeID := fmt.Sprintf("_x0000_i%d", docPrId)
	objectID := fmt.Sprintf("_%d", 1000000000+embIdx)
	oleXML := generateOLEObjectXML(shapeID, imageRelID, xlsxRelID, opts.ProgID, objectID, opts.Width, opts.Height)

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	var updated []byte
	switch opts.Position {
	case PositionBeginning:
		updated, err = insertAtBodyStart(raw, oleXML)
	case PositionEnd:
		updated, err = insertAtBodyEnd(raw, oleXML)
	case PositionAfterText:
		updated, err = insertAfterText(raw, oleXML, opts.Anchor)
	case PositionBeforeText:
		updated, err = insertBeforeText(raw, oleXML, opts.Anchor)
	default:
		return fmt.Errorf("invalid insert position: %d", opts.Position)
	}
	if err != nil {
		return fmt.Errorf("insert embedded object in document.xml: %w", err)
	}

	if err := atomicWriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// generateOLEObjectXML returns the <w:p> paragraph XML for an OLE embedded object.
// widthPt and heightPt are in typographic points; w:dxaOrig/w:dyaOrig are in twips
// (1 point = 20 twips). The DrawAspect="Icon" attribute ensures the object displays
// as a clickable icon rather than attempt to render the embedded content, matching
// MS Word's behavior for embedded file objects. Namespace prefixes are declared on
// each top-level element rather than hoisted to <w:document> for standalone-fragment
// compatibility with Word.
func generateOLEObjectXML(shapeID, imageRelID, xlsxRelID, progID, objectID string, widthPt, heightPt int) []byte {
	const tmpl = `<w:p><w:r><w:object w:dxaOrig="%d" w:dyaOrig="%d">` +
		`<v:shape id="%s" type="#_x0000_t75"` +
		` style="width:%dpt;height:%dpt" o:ole=""` +
		` xmlns:v="` + VMLNamespace + `"` +
		` xmlns:o="` + OfficeNamespace + `">` +
		`<v:imagedata r:id="%s" o:title=""` +
		` xmlns:r="` + OfficeDocumentNS + `"/>` +
		`</v:shape>` +
		`<o:OLEObject Type="Embed" ProgID="%s" ShapeID="%s"` +
		` DrawAspect="Icon" ObjectID="%s"` +
		` r:id="%s"` +
		` xmlns:o="` + OfficeNamespace + `"` +
		` xmlns:r="` + OfficeDocumentNS + `"/>` +
		`</w:object></w:r></w:p>`

	return fmt.Appendf(nil, tmpl,
		widthPt*20, heightPt*20,
		shapeID, widthPt, heightPt, imageRelID,
		progID, shapeID, objectID, xlsxRelID,
	)
}

// addDocumentRelationship adds a relationship of the given type and target to
// word/_rels/document.xml.rels and returns the new relationship ID.
func (u *Updater) addDocumentRelationship(relType, target string) (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document relationships: %w", err)
	}

	nextRelID, err := u.getNextDocumentRelId()
	if err != nil {
		return "", err
	}

	insert := fmt.Sprintf("\n  <Relationship Id=%q Type=%q Target=%q/>\n",
		nextRelID, relType, target)

	closer := []byte("</Relationships>")
	pos := bytes.LastIndex(raw, closer)
	if pos == -1 {
		return "", fmt.Errorf("invalid document.xml.rels: missing </Relationships>")
	}

	result := make([]byte, len(raw)+len(insert))
	n := copy(result, raw[:pos])
	n += copy(result[n:], []byte(insert))
	copy(result[n:], raw[pos:])

	if err := atomicWriteFile(relsPath, result, 0o644); err != nil {
		return "", fmt.Errorf("write relationships: %w", err)
	}

	return nextRelID, nil
}

// findNextEmbeddingIndex returns the next available embedding index by scanning
// word/embeddings/ for existing embeddingN.xlsx files.
func (u *Updater) findNextEmbeddingIndex() int {
	entries, err := os.ReadDir(filepath.Join(u.tempDir, "word", "embeddings"))
	if err != nil {
		return 1
	}
	maxIndex := 0
	for _, entry := range entries {
		if matches := embeddingFilePattern.FindStringSubmatch(entry.Name()); matches != nil {
			idx, err := strconv.Atoi(matches[1])
			if err == nil && idx > maxIndex {
				maxIndex = idx
			}
		}
	}
	return maxIndex + 1
}

// applyEmbedDefaults fills in zero-value fields in EmbeddedObjectOptions.
func applyEmbedDefaults(opts EmbeddedObjectOptions) EmbeddedObjectOptions {
	if opts.ProgID == "" {
		opts.ProgID = OLEProgIDExcel
	}
	if opts.Width == 0 {
		opts.Width = DefaultEmbedWidthPt
	}
	if opts.Height == 0 {
		opts.Height = DefaultEmbedHeightPt
	}
	if opts.FileName == "" {
		if opts.FilePath != "" {
			opts.FileName = filepath.Base(opts.FilePath)
		} else {
			opts.FileName = "embedded.xlsx"
		}
	}
	return opts
}

// resolveEmbedFileBytes returns the file bytes from opts.FileBytes or opts.FilePath.
func resolveEmbedFileBytes(opts EmbeddedObjectOptions) ([]byte, error) {
	if len(opts.FileBytes) > 0 {
		return opts.FileBytes, nil
	}
	if opts.FilePath == "" {
		return nil, fmt.Errorf("one of FilePath or FileBytes must be provided")
	}
	data, err := os.ReadFile(opts.FilePath)
	if err != nil {
		return nil, fmt.Errorf("read embed file %s: %w", opts.FilePath, err)
	}
	return data, nil
}

// resolveIconBytes returns icon PNG bytes from opts.IconBytes, opts.IconPath,
// or the built-in Excel icon. An unreadable IconPath silently falls back to the
// built-in so a wrong optional path never blocks document generation.
func resolveIconBytes(opts EmbeddedObjectOptions) []byte {
	if len(opts.IconBytes) > 0 {
		return opts.IconBytes
	}
	if opts.IconPath != "" {
		if data, err := os.ReadFile(opts.IconPath); err == nil {
			return data
		}
	}
	return defaultExcelIconPNG
}

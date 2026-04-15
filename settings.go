package godocx

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

const (
	settingsRelType   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	settingsPartName  = "/word/settings.xml"
	settingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"

	minimalSettingsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:updateFields w:val="1"/>
</w:settings>`
)

// ForceFieldUpdateOnOpen instructs Word and LibreOffice to recalculate all
// field codes (PAGE, NUMPAGES, DATE, TOC, cross-references, etc.) the first
// time the document is opened. This is the correct approach for programmatically
// generated DOCX files where the rendered field values are not yet available.
//
// Behaviour:
//   - If word/settings.xml already exists (e.g. from an uploaded template),
//     <w:updateFields w:val="1"/> is inserted into the existing document.
//   - If word/settings.xml does not exist, a minimal one is created and wired
//     into both [Content_Types].xml and word/_rels/document.xml.rels.
//   - Calling this method more than once is idempotent.
func (u *Updater) ForceFieldUpdateOnOpen() error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	settingsPath := filepath.Join(u.tempDir, "word", "settings.xml")

	if _, err := os.Stat(settingsPath); err == nil {
		// settings.xml exists — inject the flag into the existing file.
		return injectUpdateFields(settingsPath)
	}

	// settings.xml does not exist — create it and wire it up.
	if err := atomicWriteFile(settingsPath, []byte(minimalSettingsXML), 0o644); err != nil {
		return fmt.Errorf("write settings.xml: %w", err)
	}
	if err := u.addSettingsContentType(); err != nil {
		return fmt.Errorf("add settings content type: %w", err)
	}
	if err := u.addSettingsRelationship(); err != nil {
		return fmt.Errorf("add settings relationship: %w", err)
	}
	return nil
}

// injectUpdateFields adds <w:updateFields w:val="1"/> to an existing settings.xml.
// It is idempotent: if the element is already present nothing is written.
func injectUpdateFields(settingsPath string) error {
	raw, err := os.ReadFile(settingsPath)
	if err != nil {
		return fmt.Errorf("read settings.xml: %w", err)
	}
	content := string(raw)

	if strings.Contains(content, "w:updateFields") {
		return nil // already present
	}

	// Insert immediately before </w:settings>.
	const closeTag = "</w:settings>"
	if !strings.Contains(content, closeTag) {
		// Malformed settings.xml — append the flag and closing tag.
		content += "\n<w:updateFields w:val=\"1\"/>\n" + closeTag
	} else {
		content = strings.Replace(content, closeTag,
			"<w:updateFields w:val=\"1\"/>\n"+closeTag, 1)
	}

	return atomicWriteFile(settingsPath, []byte(content), 0o644)
}

// addSettingsContentType registers word/settings.xml in [Content_Types].xml.
func (u *Updater) addSettingsContentType() error {
	ctPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	raw, err := os.ReadFile(ctPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}
	content := string(raw)

	if strings.Contains(content, settingsPartName) {
		return nil // already registered
	}

	override := fmt.Sprintf(`<Override PartName="%s" ContentType="%s"/>`, settingsPartName, settingsContentType)
	content = strings.Replace(content, "</Types>", override+"</Types>", 1)
	return atomicWriteFile(ctPath, []byte(content), 0o644)
}

// addSettingsRelationship adds the settings relationship to word/_rels/document.xml.rels.
func (u *Updater) addSettingsRelationship() error {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read document relationships: %w", err)
	}
	content := string(raw)

	if strings.Contains(content, settingsRelType) {
		return nil // already registered
	}

	relID, err := getNextRelIDFromFile(relsPath)
	if err != nil {
		return fmt.Errorf("find next relationship id: %w", err)
	}

	newRel := fmt.Sprintf(
		`<Relationship Id="%s" Type="%s" Target="settings.xml"/>`,
		relID, settingsRelType,
	)
	content = strings.Replace(content, "</Relationships>", newRel+"</Relationships>", 1)
	return atomicWriteFile(relsPath, []byte(content), 0o644)
}

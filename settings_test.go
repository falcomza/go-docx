package godocx

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)


func TestForceFieldUpdateOnOpen_BlankDocument(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	if err := u.ForceFieldUpdateOnOpen(); err != nil {
		t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
	}

	settingsPath := filepath.Join(u.tempDir, "word", "settings.xml")
	raw, err := os.ReadFile(settingsPath)
	if err != nil {
		t.Fatalf("settings.xml not created: %v", err)
	}
	if !strings.Contains(string(raw), "w:updateFields") {
		t.Error("settings.xml missing <w:updateFields>")
	}

	// Verify content type registered
	ctRaw, _ := os.ReadFile(filepath.Join(u.tempDir, "[Content_Types].xml"))
	if !strings.Contains(string(ctRaw), "settings.xml") {
		t.Error("[Content_Types].xml missing settings.xml entry")
	}

	// Verify relationship registered
	relsRaw, _ := os.ReadFile(filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels"))
	if !strings.Contains(string(relsRaw), settingsRelType) {
		t.Error("document.xml.rels missing settings relationship")
	}
}

func TestForceFieldUpdateOnOpen_Idempotent(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	for i := 0; i < 3; i++ {
		if err := u.ForceFieldUpdateOnOpen(); err != nil {
			t.Fatalf("call %d: ForceFieldUpdateOnOpen: %v", i+1, err)
		}
	}

	// Should contain exactly one occurrence
	raw, _ := os.ReadFile(filepath.Join(u.tempDir, "word", "settings.xml"))
	count := strings.Count(string(raw), "w:updateFields")
	if count != 1 {
		t.Errorf("expected 1 occurrence of w:updateFields, got %d", count)
	}
}

func TestForceFieldUpdateOnOpen_ExistingSettings(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	// Pre-create a settings.xml without the flag (simulates an uploaded template)
	settingsPath := filepath.Join(u.tempDir, "word", "settings.xml")
	existing := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="720"/>
</w:settings>`
	if err := os.WriteFile(settingsPath, []byte(existing), 0o644); err != nil {
		t.Fatalf("write settings.xml: %v", err)
	}

	if err := u.ForceFieldUpdateOnOpen(); err != nil {
		t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
	}

	raw, _ := os.ReadFile(settingsPath)
	content := string(raw)
	if !strings.Contains(content, "w:updateFields") {
		t.Error("w:updateFields not injected into existing settings.xml")
	}
	// Original content preserved
	if !strings.Contains(content, "w:defaultTabStop") {
		t.Error("existing settings content was lost")
	}
}

func TestForceFieldUpdateOnOpen_MarksHeaderFooterFieldsDirty(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	// Add a header with a PAGE field so there is a header file to process.
	err = u.SetHeader(HeaderFooterContent{PageNumber: true}, DefaultHeaderOptions())
	if err != nil {
		t.Fatalf("SetHeader: %v", err)
	}

	if err := u.ForceFieldUpdateOnOpen(); err != nil {
		t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
	}

	// Verify the header XML carries w:dirty="true".
	entries, _ := os.ReadDir(filepath.Join(u.tempDir, "word"))
	found := false
	for _, entry := range entries {
		name := entry.Name()
		if (!strings.HasPrefix(name, "header") && !strings.HasPrefix(name, "footer")) ||
			!strings.HasSuffix(name, ".xml") {
			continue
		}
		raw, _ := os.ReadFile(filepath.Join(u.tempDir, "word", name))
		if strings.Contains(string(raw), `w:dirty="true"`) {
			found = true
		}
	}
	if !found {
		t.Error("no header/footer XML contains w:dirty=\"true\" after ForceFieldUpdateOnOpen")
	}
}

func TestForceFieldUpdateOnOpen_TemplatHeaderFieldsMarkedDirty(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	// Simulate a template-provided header that has a field without w:dirty.
	headerPath := filepath.Join(u.tempDir, "word", "header3.xml")
	templateHeader := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
</w:hdr>`
	if err := os.WriteFile(headerPath, []byte(templateHeader), 0o644); err != nil {
		t.Fatalf("write template header: %v", err)
	}

	if err := u.ForceFieldUpdateOnOpen(); err != nil {
		t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
	}

	raw, _ := os.ReadFile(headerPath)
	if !strings.Contains(string(raw), `w:dirty="true"`) {
		t.Error("template-provided header field was not marked dirty")
	}
	// Ensure the rest of the content is preserved.
	if !strings.Contains(string(raw), `w:fldCharType="separate"`) {
		t.Error("template header content was unexpectedly modified")
	}
}

func TestForceFieldUpdateOnOpen_IdempotentOnDirtyFields(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	headerPath := filepath.Join(u.tempDir, "word", "header3.xml")
	alreadyDirty := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
</w:hdr>`
	if err := os.WriteFile(headerPath, []byte(alreadyDirty), 0o644); err != nil {
		t.Fatalf("write header: %v", err)
	}

	for range 3 {
		if err := u.ForceFieldUpdateOnOpen(); err != nil {
			t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
		}
	}

	raw, _ := os.ReadFile(headerPath)
	count := strings.Count(string(raw), `w:dirty="true"`)
	if count != 1 {
		t.Errorf("expected exactly 1 w:dirty attribute, got %d", count)
	}
}

// TestInjectUpdateFields_OverridesFalseValue verifies the key bug fix: if an
// existing settings.xml already contains <w:updateFields> with a false/absent
// value, it must be replaced rather than left unchanged.
func TestInjectUpdateFields_OverridesFalseValue(t *testing.T) {
	cases := []struct {
		name  string
		input string
	}{
		{
			name: "no val attribute (defaults false)",
			input: `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:updateFields/>
</w:settings>`,
		},
		{
			name: "val=0",
			input: `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:updateFields w:val="0"/>
</w:settings>`,
		},
		{
			name: "val=false",
			input: `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:updateFields w:val="false"/>
</w:settings>`,
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			f, err := os.CreateTemp(t.TempDir(), "settings-*.xml")
			if err != nil {
				t.Fatal(err)
			}
			path := f.Name()
			f.WriteString(tc.input)
			f.Close()

			if err := injectUpdateFields(path); err != nil {
				t.Fatalf("injectUpdateFields: %v", err)
			}

			raw, _ := os.ReadFile(path)
			content := string(raw)
			if !updateFieldsEnabled(content) {
				t.Errorf("w:updateFields not set to true after injection; got:\n%s", content)
			}
			count := strings.Count(content, "w:updateFields")
			if count != 1 {
				t.Errorf("expected exactly 1 w:updateFields element, got %d", count)
			}
		})
	}
}

// TestInjectDirtyIntoFieldCodes_FldSimple verifies that <w:fldSimple> elements
// (used by LibreOffice and some older Word versions) also receive w:dirty="true".
func TestInjectDirtyIntoFieldCodes_FldSimple(t *testing.T) {
	header := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:fldSimple w:instr=" DATE \@ &quot;MMMM d, yyyy&quot; ">
<w:r><w:t>January 1, 2026</w:t></w:r>
</w:fldSimple>
</w:p>
</w:hdr>`

	f, err := os.CreateTemp(t.TempDir(), "header-*.xml")
	if err != nil {
		t.Fatal(err)
	}
	path := f.Name()
	f.WriteString(header)
	f.Close()

	if err := injectDirtyIntoFieldCodes(path); err != nil {
		t.Fatalf("injectDirtyIntoFieldCodes: %v", err)
	}

	raw, _ := os.ReadFile(path)
	if !strings.Contains(string(raw), `w:dirty="true"`) {
		t.Error("w:dirty not injected into <w:fldSimple>")
	}
}

// TestForceFieldUpdateOnOpen_EndToEnd verifies that the changes survive Save()
// by reading the output DOCX as a ZIP and inspecting the XML parts directly.
func TestForceFieldUpdateOnOpen_EndToEnd(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	if err := u.SetHeader(HeaderFooterContent{PageNumber: true, Date: true}, DefaultHeaderOptions()); err != nil {
		t.Fatalf("SetHeader: %v", err)
	}
	if err := u.ForceFieldUpdateOnOpen(); err != nil {
		t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
	}

	outPath := filepath.Join(t.TempDir(), "out.docx")
	if err := u.Save(outPath); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Read the DOCX as a ZIP and extract the XML parts we care about.
	parts := readDocxParts(t, outPath)

	settings, ok := parts["word/settings.xml"]
	if !ok {
		t.Fatal("saved DOCX missing word/settings.xml")
	}
	if !updateFieldsEnabled(settings) {
		t.Errorf("saved settings.xml does not have w:updateFields=true:\n%s", settings)
	}

	foundDirty := false
	for name, content := range parts {
		if (strings.HasPrefix(name, "word/header") || strings.HasPrefix(name, "word/footer")) &&
			strings.HasSuffix(name, ".xml") {
			if strings.Contains(content, `w:dirty="true"`) {
				foundDirty = true
			}
		}
	}
	if !foundDirty {
		t.Error("no header/footer part in saved DOCX carries w:dirty=\"true\"")
	}
}

// readDocxParts opens a DOCX file as a ZIP and returns a map of part name → content.
func readDocxParts(t *testing.T, docxPath string) map[string]string {
	t.Helper()
	r, err := zip.OpenReader(docxPath)
	if err != nil {
		t.Fatalf("open docx zip: %v", err)
	}
	defer r.Close()

	parts := make(map[string]string, len(r.File))
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open zip entry %s: %v", f.Name, err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read zip entry %s: %v", f.Name, err)
		}
		parts[f.Name] = string(data)
	}
	return parts
}

func TestForceFieldUpdateOnOpen_SavesValidDocx(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	if err := u.ForceFieldUpdateOnOpen(); err != nil {
		t.Fatalf("ForceFieldUpdateOnOpen: %v", err)
	}

	outPath := filepath.Join(t.TempDir(), "out.docx")
	if err := u.Save(outPath); err != nil {
		t.Fatalf("Save after ForceFieldUpdateOnOpen: %v", err)
	}
	if _, err := os.Stat(outPath); err != nil {
		t.Fatalf("output file not created: %v", err)
	}
}

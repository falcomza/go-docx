package godocx

import (
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

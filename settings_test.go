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

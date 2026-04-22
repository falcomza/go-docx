package godocx

import (
	"strings"
	"testing"
)

// TestNormalizeRunsInXML_SplitPlaceholder verifies that a {{PLACEHOLDER}} split
// across two consecutive runs with identical properties is merged into one run
// so that subsequent text replacement can find and substitute it.
func TestNormalizeRunsInXML_SplitPlaceholder(t *testing.T) {
	// Simulate a header XML where Word fragmented "{{REPORT_TITLE}}" across two runs.
	input := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:t>{{REPORT_</w:t></w:r><w:r><w:t>TITLE}}</w:t></w:r></w:p>` +
		`</w:hdr>`

	got := string(normalizeRunsInXML([]byte(input)))

	if strings.Count(got, "<w:r>") != 1 {
		t.Errorf("expected 1 merged run, got XML:\n%s", got)
	}
	if !strings.Contains(got, "{{REPORT_TITLE}}") {
		t.Errorf("merged text not found; got:\n%s", got)
	}
}

// TestNormalizeRunsInXML_DifferentRpr verifies that runs with different rPr are
// NOT merged, so existing per-character formatting is preserved.
func TestNormalizeRunsInXML_DifferentRpr(t *testing.T) {
	input := `<w:p>` +
		`<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>` +
		`<w:r><w:rPr><w:i/></w:rPr><w:t> italic</w:t></w:r>` +
		`</w:p>`

	got := string(normalizeRunsInXML([]byte(input)))

	if strings.Count(got, "<w:r>") != 2 {
		t.Errorf("runs with different rPr should not be merged; got:\n%s", got)
	}
}

// TestNormalizeRunsInXML_HyperlinkNotMerged verifies that runs on either side of
// a </w:hyperlink> closing tag are NOT merged. This was the root cause of the
// "w:p start tag does not match end tag of w:hyperlink" validation error.
func TestNormalizeRunsInXML_HyperlinkNotMerged(t *testing.T) {
	// Paragraph containing a hyperlink run followed by a plain run — both
	// happen to have no rPr. Must NOT be merged across the </w:hyperlink> boundary.
	input := `<w:p>` +
		`<w:hyperlink r:id="rId1">` +
		`<w:r><w:t>Click here</w:t></w:r>` +
		`</w:hyperlink>` +
		`<w:r><w:t> for details</w:t></w:r>` +
		`</w:p>`

	got := string(normalizeRunsInXML([]byte(input)))

	// The </w:hyperlink> must still be present.
	if !strings.Contains(got, "</w:hyperlink>") {
		t.Errorf("</w:hyperlink> was lost; got:\n%s", got)
	}
	// Both runs must be preserved separately — only 2 runs should exist.
	if strings.Count(got, "<w:r>") != 2 {
		t.Errorf("expected 2 runs (hyperlink and plain); got:\n%s", got)
	}
}

// TestNormalizeRunsInXML_BookmarkPreserved ensures that a <w:bookmarkEnd/>
// between two same-rPr runs is not swallowed by the merge.
func TestNormalizeRunsInXML_BookmarkPreserved(t *testing.T) {
	input := `<w:p>` +
		`<w:r><w:t>before </w:t></w:r>` +
		`<w:bookmarkEnd w:id="1"/>` +
		`<w:r><w:t>after</w:t></w:r>` +
		`</w:p>`

	got := string(normalizeRunsInXML([]byte(input)))

	if !strings.Contains(got, `<w:bookmarkEnd`) {
		t.Errorf("<w:bookmarkEnd/> was lost; got:\n%s", got)
	}
}

// TestNormalizeRunsInXML_SameRpr verifies that runs with identical rPr are merged.
func TestNormalizeRunsInXML_SameRpr(t *testing.T) {
	input := `<w:p>` +
		`<w:r><w:rPr><w:b/></w:rPr><w:t>Hello </w:t></w:r>` +
		`<w:r><w:rPr><w:b/></w:rPr><w:t>World</w:t></w:r>` +
		`</w:p>`

	got := string(normalizeRunsInXML([]byte(input)))

	if strings.Count(got, "<w:r>") != 1 {
		t.Errorf("runs with identical rPr should be merged; got:\n%s", got)
	}
	if !strings.Contains(got, "Hello World") {
		t.Errorf("merged text incorrect; got:\n%s", got)
	}
}

// TestNormalizeRunsInXML_Idempotent verifies calling normalizeRunsInXML twice
// produces the same output as calling it once.
func TestNormalizeRunsInXML_Idempotent(t *testing.T) {
	input := `<w:p>` +
		`<w:r><w:t>{{FOO_</w:t></w:r>` +
		`<w:r><w:t>BAR}}</w:t></w:r>` +
		`</w:p>`

	once := normalizeRunsInXML([]byte(input))
	twice := normalizeRunsInXML(once)

	if string(once) != string(twice) {
		t.Errorf("not idempotent:\nonce:  %s\ntwice: %s", once, twice)
	}
}

// TestNormalizeRunsInXML_PreservesSpace verifies that xml:space="preserve" is
// emitted on the merged run when the combined text has leading/trailing spaces.
func TestNormalizeRunsInXML_PreservesSpace(t *testing.T) {
	input := `<w:p>` +
		`<w:r><w:t xml:space="preserve">Page </w:t></w:r>` +
		`<w:r><w:t>1</w:t></w:r>` +
		`</w:p>`

	got := string(normalizeRunsInXML([]byte(input)))

	if !strings.Contains(got, `xml:space="preserve"`) {
		t.Errorf("expected preserve attribute on merged run; got:\n%s", got)
	}
	if !strings.Contains(got, "Page 1") {
		t.Errorf("merged text incorrect; got:\n%s", got)
	}
}

// TestMergeCompatibleRuns_ThreeWayMerge verifies three consecutive plain runs
// are collapsed into one.
func TestMergeCompatibleRuns_ThreeWayMerge(t *testing.T) {
	input := `<w:p>` +
		`<w:r><w:t>{{</w:t></w:r>` +
		`<w:r><w:t>KEY</w:t></w:r>` +
		`<w:r><w:t>}}</w:t></w:r>` +
		`</w:p>`

	got := string(mergeCompatibleRuns([]byte(input)))

	if strings.Count(got, "<w:r>") != 1 {
		t.Errorf("three plain runs should become one; got:\n%s", got)
	}
	if !strings.Contains(got, "{{KEY}}") {
		t.Errorf("merged text should be {{KEY}}; got:\n%s", got)
	}
}

// TestReplaceText_SplitHeaderPlaceholder is an integration test that verifies
// ReplaceText with InHeaders=true replaces a placeholder that was split across
// runs inside a header XML file. This exercises the full path including file I/O.
func TestReplaceText_SplitHeaderPlaceholder(t *testing.T) {
	u, err := NewBlank()
	if err != nil {
		t.Fatalf("NewBlank: %v", err)
	}
	defer u.Cleanup()

	// Write a synthetic header1.xml with a split placeholder directly to tempDir.
	headerPath := u.tempDir + "/word/header1.xml"
	headerXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:t>{{REPORT_</w:t></w:r><w:r><w:t>TITLE}}</w:t></w:r></w:p>` +
		`</w:hdr>`

	if err := atomicWriteFile(headerPath, []byte(headerXML), 0o644); err != nil {
		t.Fatalf("write header1.xml: %v", err)
	}

	opts := DefaultReplaceOptions()
	opts.InHeaders = true
	opts.InParagraphs = false
	opts.InTables = false

	n, err := u.ReplaceText("{{REPORT_TITLE}}", "Q1 Report", opts)
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if n != 1 {
		t.Errorf("expected 1 replacement, got %d", n)
	}
}

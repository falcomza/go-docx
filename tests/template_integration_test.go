package docxupdater_test

import (
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

// TestTemplateIntegration tests all new features using the real docx_template.docx
// This demonstrates a realistic workflow combining multiple operations
func TestTemplateIntegration(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/template_integration_test.docx"

	// Ensure template exists
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Fatalf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	// Open template
	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Test 1: Read Operations - Extract existing text
	t.Log("Testing read operations...")
	text, err := u.GetText()
	if err != nil {
		t.Fatalf("Failed to get text: %v", err)
	}
	if text == "" {
		t.Error("Expected non-empty text from template")
	}
	t.Logf("Extracted %d characters from template", len(text))

	// Test 2: Find Text - Search for specific patterns
	t.Log("Testing find text...")
	paragraphText, err := u.GetParagraphText()
	if err != nil {
		t.Fatalf("Failed to get paragraph text: %v", err)
	}
	if len(paragraphText) == 0 {
		t.Error("Expected at least one paragraph")
	}
	t.Logf("Found %d paragraphs in template", len(paragraphText))

	// Test 3: Text Replacement - Replace placeholders
	t.Log("Testing text replacement...")
	replaceOpts := docxupdater.DefaultReplaceOptions()
	replaceOpts.MatchCase = false

	// Replace any date patterns (example)
	count, err := u.ReplaceText("2024", "2026", replaceOpts)
	if err != nil {
		t.Fatalf("Failed to replace text: %v", err)
	}
	t.Logf("Replaced %d occurrences of '2024' with '2026'", count)

	// Test 4: Regex Replacement - Replace email patterns
	t.Log("Testing regex replacement...")
	emailPattern := regexp.MustCompile(`\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b`)
	count, err = u.ReplaceTextRegex(emailPattern, "contact@example.com", replaceOpts)
	if err != nil {
		t.Fatalf("Failed to replace text with regex: %v", err)
	}
	t.Logf("Replaced %d email addresses", count)

	// Test 5: Insert Hyperlink
	t.Log("Testing hyperlink insertion...")
	hyperlinkOpts := docxupdater.DefaultHyperlinkOptions()
	hyperlinkOpts.Position = docxupdater.PositionEnd
	hyperlinkOpts.Tooltip = "Visit our website for more information"

	err = u.InsertHyperlink("Click here for more info", "https://github.com/falcomza/docx-update", hyperlinkOpts)
	if err != nil {
		t.Fatalf("Failed to insert hyperlink: %v", err)
	}
	t.Log("Successfully inserted hyperlink")

	// Test 6: Set Header
	t.Log("Testing header creation...")
	headerContent := docxupdater.HeaderFooterContent{
		LeftText:   "DOCX-Update Library",
		CenterText: "Integration Test Document",
		RightText:  "February 2026",
	}
	headerOpts := docxupdater.DefaultHeaderOptions()

	err = u.SetHeader(headerContent, headerOpts)
	if err != nil {
		t.Fatalf("Failed to set header: %v", err)
	}
	t.Log("Successfully created header")

	// Test 7: Set Footer with Page Numbers
	t.Log("Testing footer creation...")
	footerContent := docxupdater.HeaderFooterContent{
		LeftText:         "Confidential",
		CenterText:       "Page ",
		PageNumber:       true,
		PageNumberFormat: "X of Y",
		RightText:        "Generated: 2026-02-16",
	}
	footerOpts := docxupdater.DefaultFooterOptions()

	err = u.SetFooter(footerContent, footerOpts)
	if err != nil {
		t.Fatalf("Failed to set footer: %v", err)
	}
	t.Log("Successfully created footer")

	// Save the modified document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify output file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}

	info, err := os.Stat(outputPath)
	if err != nil {
		t.Fatalf("Failed to stat output file: %v", err)
	}
	if info.Size() == 0 {
		t.Error("Output file is empty")
	}

	t.Logf("Successfully created output file: %s (size: %d bytes)", outputPath, info.Size())
}

// TestTemplateReplaceMultiple tests replacing multiple different strings
func TestTemplateReplaceMultiple(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/template_replace_multiple_test.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	opts := docxupdater.DefaultReplaceOptions()
	opts.MatchCase = false

	// Perform multiple replacements
	replacements := map[string]string{
		"Document": "Report",
		"Test":     "Production",
		"Example":  "Sample",
	}

	totalCount := 0
	for old, new := range replacements {
		count, err := u.ReplaceText(old, new, opts)
		if err != nil {
			t.Fatalf("Failed to replace '%s' with '%s': %v", old, new, err)
		}
		t.Logf("Replaced %d occurrences of '%s' with '%s'", count, old, new)
		totalCount += count
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	t.Logf("Total replacements: %d", totalCount)
}

// TestTemplateFindAll tests finding all occurrences of patterns
func TestTemplateFindAll(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Find all numbers
	opts := docxupdater.DefaultFindOptions()
	opts.UseRegex = true

	matches, err := u.FindText(`\d+`, opts)
	if err != nil {
		t.Fatalf("Failed to find text: %v", err)
	}

	t.Logf("Found %d numeric values in document", len(matches))

	// Display first few matches
	for i, match := range matches {
		if i >= 5 {
			t.Log("... and more")
			break
		}
		context := match.Before + match.Text + match.After
		t.Logf("Match %d: '%s' at paragraph %d (context: ...%s...)",
			i+1, match.Text, match.Paragraph, context)
	}
}

// TestTemplateExtractTables tests extracting table content
func TestTemplateExtractTables(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Extract table content
	tableText, err := u.GetTableText()
	if err != nil {
		t.Fatalf("Failed to get table text: %v", err)
	}

	t.Logf("Found %d tables in document", len(tableText))

	for i, table := range tableText {
		t.Logf("Table %d: %d rows", i+1, len(table))
		if len(table) > 0 {
			// Show first row
			t.Logf("  First row: %s", strings.Join(table[0], " | "))
		}
	}
}

// TestTemplateErrorHandling tests error handling with invalid operations
func TestTemplateErrorHandling(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Test invalid hyperlink URL
	err = u.InsertHyperlink("Bad Link", "not-a-valid-url", docxupdater.DefaultHyperlinkOptions())
	if err == nil {
		t.Error("Expected error for invalid URL, got nil")
	}

	// Check it's a DocxError
	docxErr, ok := err.(*docxupdater.DocxError)
	if !ok {
		t.Errorf("Expected DocxError type, got %T", err)
	} else {
		t.Logf("Correctly returned DocxError with code: %s", docxErr.Code)
		if docxErr.Code != docxupdater.ErrCodeInvalidURL {
			t.Errorf("Expected error code %s, got %s", docxupdater.ErrCodeInvalidURL, docxErr.Code)
		}
	}
}

// TestTemplateCompleteWorkflow demonstrates a complete realistic workflow
func TestTemplateCompleteWorkflow(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/template_complete_workflow.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	t.Log("Step 1: Opening template...")
	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	t.Log("Step 2: Extracting current content...")
	text, err := u.GetText()
	if err != nil {
		t.Fatalf("Failed to extract text: %v", err)
	}
	t.Logf("  Document has %d characters", len(text))

	t.Log("Step 3: Performing content replacements...")
	opts := docxupdater.DefaultReplaceOptions()
	opts.MatchCase = false

	// Replace company name
	count1, _ := u.ReplaceText("ACME Corp", "TechVenture Inc", opts)
	// Update dates
	count2, _ := u.ReplaceText("2023", "2026", opts)
	// Update version numbers
	pattern := regexp.MustCompile(`v\d+\.\d+`)
	count3, _ := u.ReplaceTextRegex(pattern, "v2.0", opts)

	t.Logf("  Made %d content replacements", count1+count2+count3)

	t.Log("Step 4: Adding professional header...")
	headerContent := docxupdater.HeaderFooterContent{
		LeftText:   "TechVenture Inc",
		CenterText: "Annual Report 2026",
		RightText:  "CONFIDENTIAL",
	}
	if err := u.SetHeader(headerContent, docxupdater.DefaultHeaderOptions()); err != nil {
		t.Fatalf("Failed to set header: %v", err)
	}

	t.Log("Step 5: Adding footer with page numbers...")
	footerContent := docxupdater.HeaderFooterContent{
		LeftText:         "© 2026 TechVenture Inc",
		CenterText:       "Page ",
		PageNumber:       true,
		PageNumberFormat: "X of Y",
		RightText:        "www.techventure.com",
	}
	if err := u.SetFooter(footerContent, docxupdater.DefaultFooterOptions()); err != nil {
		t.Fatalf("Failed to set footer: %v", err)
	}

	t.Log("Step 6: Adding hyperlinks...")
	linkOpts := docxupdater.DefaultHyperlinkOptions()
	linkOpts.Position = docxupdater.PositionEnd
	linkOpts.Color = "0563C1" // Standard hyperlink blue

	if err := u.InsertHyperlink("Visit our website", "https://techventure.com", linkOpts); err != nil {
		t.Fatalf("Failed to insert hyperlink: %v", err)
	}

	if err := u.InsertHyperlink("Contact us", "mailto:info@techventure.com", linkOpts); err != nil {
		t.Fatalf("Failed to insert email link: %v", err)
	}

	t.Log("Step 7: Saving final document...")
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the output
	info, err := os.Stat(outputPath)
	if err != nil {
		t.Fatalf("Failed to stat output file: %v", err)
	}

	t.Logf("✓ Successfully created document: %s", outputPath)
	t.Logf("  File size: %d bytes", info.Size())
	t.Log("  All operations completed successfully!")
}

package docxupdater

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"os"
	"strings"
	"testing"
)

// TestWordCompatibility_XMLFormatting tests that generated DOCX files have proper XML formatting
// that Microsoft Word can open without corruption errors
func TestWordCompatibility_XMLFormatting(t *testing.T) {
	// Create a simple document with a chart
	updater, err := New("templates/docx_template.docx")
	if err != nil {
		t.Skipf("Template not found (expected in CI): %v", err)
		return
	}
	defer updater.Cleanup()

	// Add a chart
	err = updater.InsertChart(ChartOptions{
		Position:          PositionEnd,
		ChartKind:         ChartKindColumn,
		Title:             "Test Chart",
		CategoryAxisTitle: "Categories",
		ValueAxisTitle:    "Values",
		Categories:        []string{"A", "B", "C"},
		Series: []SeriesData{
			{Name: "Series 1", Values: []float64{10, 20, 30}},
			{Name: "Series 2", Values: []float64{15, 25, 35}},
		},
		ShowLegend:     true,
		LegendPosition: "r",
	})
	if err != nil {
		t.Fatalf("Failed to insert chart: %v", err)
	}

	// Save the document
	outputPath := "outputs/test_word_compatibility.docx"
	if err := updater.Save(outputPath); err != nil {
		t.Fatalf("Failed to save: %v", err)
	}
	defer os.Remove(outputPath)

	// Now validate the XML structure
	t.Run("ValidateXMLStructure", func(t *testing.T) {
		validateDocxXMLStructure(t, outputPath)
	})
}

// validateDocxXMLStructure checks that all XML files in the DOCX have proper formatting
func validateDocxXMLStructure(t *testing.T, docxPath string) {
	r, err := zip.OpenReader(docxPath)
	if err != nil {
		t.Fatalf("Failed to open DOCX: %v", err)
	}
	defer r.Close()

	for _, f := range r.File {
		// Check XML and rels files
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}

		t.Run(f.Name, func(t *testing.T) {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("Failed to open file: %v", err)
			}
			defer rc.Close()

			content, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("Failed to read file: %v", err)
			}

			// 1. Check XML is valid
			var doc any
			if err := xml.Unmarshal(content, &doc); err != nil {
				t.Errorf("❌ Invalid XML: %v", err)
				t.Logf("First 500 chars: %s", string(content[:min(500, len(content))]))
			} else {
				t.Logf("✓ Valid XML")
			}

			contentStr := string(content)

			// 2. Check for newline after XML declaration
			if strings.Contains(contentStr, "<?xml") {
				declEnd := strings.Index(contentStr, "?>")
				if declEnd != -1 {
					afterDecl := declEnd + 2
					if afterDecl < len(contentStr) {
						nextChar := contentStr[afterDecl]
						if nextChar != '\n' && nextChar != '\r' {
							t.Errorf("❌ Missing newline after XML declaration")
							t.Logf("Characters after declaration: %q", contentStr[afterDecl:min(afterDecl+50, len(contentStr))])
						} else {
							t.Logf("✓ Has newline after XML declaration")
						}
					}
				}
			}

			// 3. Check for required namespaces in chart files
			if strings.Contains(f.Name, "charts/chart") && strings.HasSuffix(f.Name, ".xml") {
				requiredNS := []string{
					"http://schemas.openxmlformats.org/drawingml/2006/chart",
					"http://schemas.openxmlformats.org/drawingml/2006/main",
					"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
				}

				recommendedNS := []string{
					"http://schemas.microsoft.com/office/drawing/2015/06/chart",
				}

				for _, ns := range requiredNS {
					if !strings.Contains(contentStr, ns) {
						t.Errorf("❌ Missing required namespace: %s", ns)
					}
				}

				foundRecommended := 0
				for _, ns := range recommendedNS {
					if strings.Contains(contentStr, ns) {
						foundRecommended++
					}
				}

				if foundRecommended > 0 {
					t.Logf("✓ Has recommended namespace(s)")
				} else {
					t.Logf("⚠️  Missing recommended namespace (may cause issues in Word): %s", recommendedNS[0])
				}

				// Check for chart properties
				chartProps := []string{"<c:date1904", "<c:lang", "<c:roundedCorners"}
				foundProps := 0
				for _, prop := range chartProps {
					if strings.Contains(contentStr, prop) {
						foundProps++
					}
				}
				if foundProps > 0 {
					t.Logf("✓ Has chart properties (%d/%d)", foundProps, len(chartProps))
				}
			}

			// 4. Check for control characters (except tab, newline, carriage return)
			for i, ch := range content {
				if ch < 0x20 && ch != 0x09 && ch != 0x0A && ch != 0x0D {
					t.Errorf("❌ Invalid control character 0x%02X at byte %d", ch, i)
					break
				}
			}
		})
	}
}

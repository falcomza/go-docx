package docxupdater_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update/src"
)

// Verifies that when updating a copied chart with fewer series than the template,
// extra <c:ser> elements are removed from chartN.xml
func TestUpdateDropsExtraSeries(t *testing.T) {
	tpl := "../templates/docx_output_10_rows.docx"
	if _, err := os.Stat(tpl); err != nil {
		t.Skip("template not present: " + tpl)
	}

	u, err := docxupdater.New(tpl)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	// Copy chart 1 (do not depend on text anchors)
	newIdx, err := u.CopyChart(1, "")
	if err != nil {
		t.Fatalf("CopyChart: %v", err)
	}

	// Update with a single series, while the template has >= 2
	data := docxupdater.ChartData{
		Categories: []string{"A", "B", "C"},
		Series: []docxupdater.SeriesData{{
			Name:   "Only",
			Values: []float64{1, 2, 3},
		}},
	}
	if err := u.UpdateChart(newIdx, data); err != nil {
		t.Fatalf("UpdateChart: %v", err)
	}

	// Inspect chartN.xml to ensure there is only one <ser>
	chartPath := filepath.Join(u.TempDir(), "word", "charts",
		"chart"+strconvItoa(newIdx)+".xml")
	b, err := os.ReadFile(chartPath)
	if err != nil {
		t.Fatalf("read chart xml: %v", err)
	}
	count := strings.Count(string(b), "<ser")
	if count != 1 {
		t.Fatalf("expected exactly 1 <ser>, got %d", count)
	}
}

// minimal int->string to avoid importing strconv in test
func strconvItoa(i int) string {
	return string('0' + rune(i))
}

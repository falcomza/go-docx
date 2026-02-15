package main

import (
	"log"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Use the existing 10-row output as a clean test base
	u, err := updater.New("templates/docx_output_10_rows.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Just copy the chart once
	newIndex, err := u.CopyChart(1, "Contents")
	if err != nil {
		log.Fatalf("Failed to copy chart: %v", err)
	}

	// Update with simple data
	data := updater.ChartData{
		Categories: []string{"A", "B", "C"},
		Series: []updater.SeriesData{
			{Name: "Test Series", Values: []float64{10, 20, 30}},
		},
		ChartTitle: "Test Chart",
	}

	if err := u.UpdateChart(newIndex, data); err != nil {
		log.Fatal(err)
	}

	if err := u.Save("outputs/test_simple_output.docx"); err != nil {
		log.Fatal(err)
	}

	log.Printf("Success! Created chart %d", newIndex)
}

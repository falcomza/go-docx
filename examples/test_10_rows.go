package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	docxupdater "github.com/falcomza/docx-update/src"
)

func main() {
	// Get the current working directory
	wd, err := os.Getwd()
	if err != nil {
		log.Fatalf("Failed to get working directory: %v", err)
	}

	inputPath := filepath.Join(wd, "templates", "docx_template.docx")
	outputPath := filepath.Join(wd, "templates", "docx_output_10_rows.docx")

	fmt.Printf("Input file: %s\n", inputPath)
	fmt.Printf("Output file: %s\n", outputPath)

	// Create updater
	updater, err := docxupdater.New(inputPath)
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Cleanup()

	// Create chart data with 10 rows
	data := docxupdater.ChartData{
		Categories: []string{
			"Device 1",
			"Device 2",
			"Device 3",
			"Device 4",
			"Device 5",
			"Device 6",
			"Device 7",
			"Device 8",
			"Device 9",
			"Device 10",
		},
		Series: []docxupdater.SeriesData{
			{
				Name: "Critical",
				Values: []float64{
					10.5, 20.3, 15.8, 25.2, 30.1,
					18.7, 22.9, 27.4, 12.6, 35.8,
				},
			},
			{
				Name: "Non-critical",
				Values: []float64{
					5.2, 15.7, 8.3, 12.9, 22.4,
					9.8, 14.5, 19.1, 6.7, 28.3,
				},
			},
		},
		// Custom titles
		ChartTitle:        "Alarm Statistics Report",
		CategoryAxisTitle: "Network Devices",
		ValueAxisTitle:    "Alarm Count",
	}

	// Update the first chart
	fmt.Println("Updating chart with 10 data rows...")
	if err := updater.UpdateChart(1, data); err != nil {
		log.Fatalf("Failed to update chart: %v", err)
	}

	// Save the output
	fmt.Println("Saving updated document...")
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Printf("✓ Successfully created document with 10 data rows: %s\n", outputPath)
}

package main

import (
	"fmt"
	"log"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Open template DOCX - should contain one example chart and marker text for each subsystem
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Define subsystems and their performance data
	type SubsystemData struct {
		Name   string
		Marker string // Text in document after which to insert the chart
		Data   updater.ChartData
	}

	subsystems := []SubsystemData{
		{
			Name:   "Authentication Service",
			Marker: "Contents", // First chart uses template chart, so marker not used
			Data: updater.ChartData{
				Categories: []string{"Jan", "Feb", "Mar", "Apr"},
				Series: []updater.SeriesData{
					{Name: "Requests", Values: []float64{1200, 1450, 1380, 1520}},
					{Name: "Errors", Values: []float64{15, 12, 18, 10}},
				},
				ChartTitle:        "Authentication Service Performance",
				CategoryAxisTitle: "Month",
				ValueAxisTitle:    "Count",
			},
		},
		{
			Name:   "Database Service",
			Marker: "Contents", // Insert after this text in document
			Data: updater.ChartData{
				Categories: []string{"Jan", "Feb", "Mar", "Apr"},
				Series: []updater.SeriesData{
					{Name: "Queries", Values: []float64{5600, 6200, 5900, 6400}},
					{Name: "Cache Hits", Values: []float64{4200, 4800, 4500, 5100}},
				},
				ChartTitle:        "Database Service Performance",
				CategoryAxisTitle: "Month",
				ValueAxisTitle:    "Count",
			},
		},
		{
			Name:   "API Gateway",
			Marker: "Contents",
			Data: updater.ChartData{
				Categories: []string{"Jan", "Feb", "Mar", "Apr"},
				Series: []updater.SeriesData{
					{Name: "Total Requests", Values: []float64{8900, 9500, 9200, 10100}},
					{Name: "Throttled", Values: []float64{120, 98, 105, 89}},
				},
				ChartTitle:        "API Gateway Performance",
				CategoryAxisTitle: "Month",
				ValueAxisTitle:    "Count",
			},
		},
	}

	// Process each subsystem
	for i, subsystem := range subsystems {
		var chartIndex int

		if i == 0 {
			// Use the template chart for the first subsystem
			chartIndex = 1
			fmt.Printf("Using template chart (index 1) for %s\n", subsystem.Name)
		} else {
			// Copy the template chart for subsequent subsystems
			chartIndex, err = u.CopyChart(1, subsystem.Marker)
			if err != nil {
				log.Fatalf("Failed to copy chart for %s: %v", subsystem.Name, err)
			}
			fmt.Printf("Created chart %d for %s\n", chartIndex, subsystem.Name)
		}

		// Update the chart with subsystem-specific data
		if err := u.UpdateChart(chartIndex, subsystem.Data); err != nil {
			log.Fatalf("Failed to update chart %d for %s: %v", chartIndex, subsystem.Name, err)
		}
		fmt.Printf("Updated chart %d with data for %s\n", chartIndex, subsystem.Name)
	}

	// Save the final document
	outputPath := "outputs/docx_multi_subsystem_output.docx"
	if err := u.Save(outputPath); err != nil {
		log.Fatal(err)
	}

	fmt.Printf("\nSuccessfully generated document with %d subsystem charts: %s\n", len(subsystems), outputPath)
}

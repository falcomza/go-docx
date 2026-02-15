package main

import (
	"fmt"
	"log"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Use template that already has 3 charts inserted
	// (created by Insert → Chart three times in LibreOffice/Word)
	u, err := updater.New("templates/docx_output_10_rows.docx") // Has 1 chart
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Define data for multiple subsystems
	subsystems := []struct {
		Name string
		Data updater.ChartData
	}{
		{
			Name: "Authentication Service",
			Data: updater.ChartData{
				Categories: []string{"Jan", "Feb", "Mar", "Apr"},
				Series: []updater.SeriesData{
					{Name: "Requests", Values: []float64{1200, 1450, 1380, 1520}},
					{Name: "Errors", Values: []float64{15, 12, 18, 10}},
				},
				ChartTitle:        "Authentication Service",
				CategoryAxisTitle: "Month",
				ValueAxisTitle:    "Count",
			},
		},
	}

	// Update chart 1 with first subsystem data
	if err := u.UpdateChart(1, subsystems[0].Data); err != nil {
		log.Fatal(err)
	}

	fmt.Println("✓ Updated chart 1")

	// TODO: If template had charts 2 and 3, update them here:
	// if err := u.UpdateChart(2, subsystems[1].Data); err != nil { log.Fatal(err) }
	// if err := u.UpdateChart(3, subsystems[2].Data); err != nil { log.Fatal(err) }

	if err := u.Save("outputs/output_template_approach.docx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("✓ Saved outputs/output_template_approach.docx")
}

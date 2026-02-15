package main

import (
	"log"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	u, err := updater.New("templates/docx_output_10_rows.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Just copy the chart, DON'T update it
	newIndex, err := u.CopyChart(1, "Contents")
	if err != nil {
		log.Fatalf("Failed to copy chart: %v", err)
	}

	if err := u.Save("outputs/test_copy_no_update_output.docx"); err != nil {
		log.Fatal(err)
	}

	log.Printf("Success! Created chart %d (without updating)", newIndex)
}

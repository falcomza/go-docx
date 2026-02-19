package main

import (
	"fmt"
	"log"

	docx "github.com/falcomza/docx-update"
)

func main() {
	templatePath := "templates/docx_template.docx"
	outputPath := "outputs/test_addtext_only.docx"

	updater, err := docx.New(templatePath)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Just add text, NO replacements
	err = updater.AddText("Test paragraph 1", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	err = updater.AddText("Test paragraph 2", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Println("Document created successfully!")
}

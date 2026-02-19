package main

import (
	"fmt"
	"log"

	docx "github.com/falcomza/docx-update"
)

func main() {
	templatePath := "templates/docx_template.docx"
	outputPath := "outputs/test_replace_only.docx"

	updater, err := docx.New(templatePath)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Just do replacements, NO AddText
	_, err = updater.ReplaceText("Subsystem", "TestSystem", docx.ReplaceOptions{
		InParagraphs: true,
	})
	if err != nil {
		log.Fatalf("Failed to replace text: %v", err)
	}

	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Println("Document created successfully!")
}

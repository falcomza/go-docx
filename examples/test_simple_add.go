package main

import (
	"fmt"
	"log"
	"os"

	docx "github.com/falcomza/docx-update"
)

func main() {
	// Create a simpler test to isolate the issue
	templatePath := "templates/docx_template.docx"
	outputPath := "outputs/test_simple.docx"

	updater, err := docx.New(templatePath)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Just add simple text with placeholders
	err = updater.AddText("Company: {{COMPANY}}", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	err = updater.AddText("Date: {{DATE}}", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add text: %v", err)
	}

	// Now do replacements
	_, err = updater.ReplaceText("{{COMPANY}}", "Acme Corporation", docx.ReplaceOptions{
		InParagraphs: true,
	})
	if err != nil {
		log.Fatalf("Failed to replace text: %v", err)
	}

	_, err = updater.ReplaceText("{{DATE}}", "2026-02-19", docx.ReplaceOptions{
		InParagraphs: true,
	})
	if err != nil {
		log.Fatalf("Failed to replace text: %v", err)
	}

	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Println("Document created successfully!")

	// Now verify it
	if _, err := os.Stat(outputPath); err != nil {
		log.Fatalf("Output file doesn't exist: %v", err)
	}

	fmt.Println("File exists, checking with debug tool...")
}

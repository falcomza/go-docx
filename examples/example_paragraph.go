package main

import (
	"log"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Open a DOCX file
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Add a title at the beginning
	if err := u.AddHeading(1, "Performance Report", updater.PositionBeginning); err != nil {
		log.Fatal(err)
	}

	// Add a subtitle
	if err := u.AddHeading(2, "Executive Summary", updater.PositionEnd); err != nil {
		log.Fatal(err)
	}

	// Add normal paragraphs
	if err := u.AddText("This report presents the performance metrics for Q1 2026.", updater.PositionEnd); err != nil {
		log.Fatal(err)
	}

	// Add a paragraph with custom formatting
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:      "Key Findings:",
		Style:     updater.StyleNormal,
		Position:  updater.PositionEnd,
		Bold:      true,
		Underline: true,
	}); err != nil {
		log.Fatal(err)
	}

	// Add multiple paragraphs at once
	paragraphs := []updater.ParagraphOptions{
		{
			Text:     "Detailed Analysis",
			Style:    updater.StyleHeading2,
			Position: updater.PositionEnd,
		},
		{
			Text:     "The system showed improved performance across all metrics.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Response time decreased by 25%",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Italic:   true,
		},
		{
			Text:     "Error rate reduced to 0.1%",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Italic:   true,
		},
	}

	if err := u.InsertParagraphs(paragraphs); err != nil {
		log.Fatal(err)
	}

	// Add a conclusion with special formatting
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Note: All metrics are measured against Q4 2025 baseline.",
		Style:    updater.StyleQuote,
		Position: updater.PositionEnd,
		Italic:   true,
	}); err != nil {
		log.Fatal(err)
	}

	// Save the document
	if err := u.Save("outputs/paragraph_example_output.docx"); err != nil {
		log.Fatal(err)
	}

	log.Println("✓ Successfully created document with paragraphs")
}

package main

import (
	"log"

	updater "github.com/falcomza/docx-updater/src"
)

func main() {
	// Open a DOCX file
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Example 1: Create a multi-page document with page breaks
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Chapter 1: Introduction",
		Style:    updater.StyleHeading1,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This is the introduction section with important information.",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Insert page break to start Chapter 2 on a new page
	if err := u.InsertPageBreak(updater.BreakOptions{
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Chapter 2: Methodology",
		Style:    updater.StyleHeading1,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This section describes the methodology used in the research.",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Insert another page break
	if err := u.InsertPageBreak(updater.BreakOptions{
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Chapter 3: Results",
		Style:    updater.StyleHeading1,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 2: Insert page break after specific text
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "End of main content. See appendix for additional details.",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Insert page break after "main content" to separate appendix
	if err := u.InsertPageBreak(updater.BreakOptions{
		Position: updater.PositionAfterText,
		Anchor:   "End of main content",
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Appendix A: Additional Data",
		Style:    updater.StyleHeading1,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 3: Section breaks for different formatting
	// Insert a section break to allow different page orientation/margins
	if err := u.InsertSectionBreak(updater.BreakOptions{
		Position:    updater.PositionEnd,
		SectionType: updater.SectionBreakNextPage,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Technical Specifications (New Section)",
		Style:    updater.StyleHeading1,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This section uses a different page layout configuration.",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 4: Continuous section break (same page, different columns)
	if err := u.InsertSectionBreak(updater.BreakOptions{
		Position:    updater.PositionEnd,
		SectionType: updater.SectionBreakContinuous,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Multi-column section starts here (continuous break).",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 5: Even/Odd page section breaks (for double-sided printing)
	if err := u.InsertSectionBreak(updater.BreakOptions{
		Position:    updater.PositionEnd,
		SectionType: updater.SectionBreakEvenPage,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This content starts on the next even page.",
		Style:    updater.StyleHeading2,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertSectionBreak(updater.BreakOptions{
		Position:    updater.PositionEnd,
		SectionType: updater.SectionBreakOddPage,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This content starts on the next odd page.",
		Style:    updater.StyleHeading2,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 6: Mix page breaks with other content
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "References",
		Style:    updater.StyleHeading1,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Insert page break before references
	if err := u.InsertPageBreak(updater.BreakOptions{
		Position: updater.PositionBeforeText,
		Anchor:   "References",
	}); err != nil {
		log.Fatal(err)
	}

	// Save the document
	if err := u.Save("output/document_with_breaks.docx"); err != nil {
		log.Fatal(err)
	}

	log.Println("Document with page and section breaks created successfully!")
}

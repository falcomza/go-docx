package main

import (
	"fmt"
	"log"

	docxupdater "github.com/falcomza/docx-update/src"
)

func main() {
	// Create updater from template or blank document
	u, err := docxupdater.New("template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer u.Cleanup()

	fmt.Println("Bookmark Examples")
	fmt.Println("=================")

	// Example 1: Create an empty bookmark (position marker)
	fmt.Println("\n1. Creating empty bookmarks as position markers...")
	emptyOpts := docxupdater.DefaultBookmarkOptions()
	emptyOpts.Position = docxupdater.PositionEnd

	if err := u.CreateBookmark("section_marker", emptyOpts); err != nil {
		log.Printf("Warning: Failed to create empty bookmark: %v", err)
	}

	// Example 2: Create bookmark with text content
	fmt.Println("2. Creating bookmarks with text content...")
	textOpts := docxupdater.DefaultBookmarkOptions()
	textOpts.Position = docxupdater.PositionEnd
	textOpts.Style = docxupdater.StyleHeading1

	if err := u.CreateBookmarkWithText("executive_summary", "Executive Summary", textOpts); err != nil {
		log.Printf("Warning: Failed to create bookmark with text: %v", err)
	}

	// Add some content to the section
	if err := u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This document provides a comprehensive overview of our findings and recommendations.",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	}); err != nil {
		log.Printf("Warning: Failed to insert paragraph: %v", err)
	}

	// Example 3: Create more bookmarked sections
	fmt.Println("3. Creating multiple bookmarked sections...")

	sections := map[string]string{
		"introduction":    "Introduction",
		"methodology":     "Methodology",
		"results":         "Results and Analysis",
		"conclusion":      "Conclusion",
		"recommendations": "Recommendations",
	}

	for bookmarkName, title := range sections {
		sectionOpts := docxupdater.DefaultBookmarkOptions()
		sectionOpts.Position = docxupdater.PositionEnd
		sectionOpts.Style = docxupdater.StyleHeading2

		if err := u.CreateBookmarkWithText(bookmarkName, title, sectionOpts); err != nil {
			log.Printf("Warning: Failed to create bookmark '%s': %v", bookmarkName, err)
			continue
		}

		// Add sample content to each section
		content := fmt.Sprintf("Content for the %s section goes here.", title)
		if err := u.InsertParagraph(docxupdater.ParagraphOptions{
			Text:     content,
			Style:    docxupdater.StyleNormal,
			Position: docxupdater.PositionEnd,
		}); err != nil {
			log.Printf("Warning: Failed to insert paragraph: %v", err)
		}
	}

	// Example 4: Create table of contents with internal links
	fmt.Println("4. Creating table of contents with internal links...")

	// Insert TOC heading at the beginning
	tocHeadingOpts := docxupdater.ParagraphOptions{
		Text:     "Table of Contents",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionBeginning,
	}
	if err := u.InsertParagraph(tocHeadingOpts); err != nil {
		log.Printf("Warning: Failed to insert TOC heading: %v", err)
	}

	// Create links to each bookmarked section
	tocItems := []struct {
		text     string
		bookmark string
	}{
		{"1. Executive Summary", "executive_summary"},
		{"2. Introduction", "introduction"},
		{"3. Methodology", "methodology"},
		{"4. Results and Analysis", "results"},
		{"5. Conclusion", "conclusion"},
		{"6. Recommendations", "recommendations"},
	}

	for _, item := range tocItems {
		linkOpts := docxupdater.DefaultHyperlinkOptions()
		linkOpts.Position = docxupdater.PositionAfterText
		linkOpts.Anchor = "Table of Contents"
		linkOpts.Color = "0563C1" // Word blue
		linkOpts.Underline = true

		if err := u.InsertInternalLink(item.text, item.bookmark, linkOpts); err != nil {
			log.Printf("Warning: Failed to create internal link for '%s': %v", item.text, err)
		}
	}

	// Example 5: Wrap existing text in a bookmark
	fmt.Println("5. Wrapping existing text in bookmarks...")

	// First add some text
	if err := u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This is a key finding that we want to reference later.",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	}); err != nil {
		log.Printf("Warning: Failed to insert paragraph: %v", err)
	}

	// Now wrap part of it in a bookmark
	if err := u.WrapTextInBookmark("key_finding", "key finding"); err != nil {
		log.Printf("Warning: Failed to wrap text in bookmark: %v", err)
	}

	// Create a reference to the key finding
	refLinkOpts := docxupdater.DefaultHyperlinkOptions()
	refLinkOpts.Position = docxupdater.PositionEnd
	if err := u.InsertInternalLink("See the key finding above", "key_finding", refLinkOpts); err != nil {
		log.Printf("Warning: Failed to create reference link: %v", err)
	}

	// Example 6: Create bookmarks at specific positions
	fmt.Println("6. Creating bookmarks at specific positions...")

	// Add anchor text
	if err := u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This is the middle section of the document.",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	}); err != nil {
		log.Printf("Warning: Failed to insert paragraph: %v", err)
	}

	// Insert bookmark after specific text
	afterOpts := docxupdater.DefaultBookmarkOptions()
	afterOpts.Position = docxupdater.PositionAfterText
	afterOpts.Anchor = "middle section"
	if err := u.CreateBookmarkWithText("after_middle", "Content After Middle", afterOpts); err != nil {
		log.Printf("Warning: Failed to create bookmark after text: %v", err)
	}

	// Example 7: Valid and invalid bookmark names
	fmt.Println("7. Demonstrating bookmark name validation...")

	validNames := []string{
		"valid_bookmark",
		"ValidBookmark123",
		"My_Bookmark_Name",
	}

	invalidNames := []string{
		"invalid bookmark", // contains space
		"1invalid",         // starts with digit
		"invalid-name",     // contains hyphen
		"_Tocinvalid",      // reserved prefix
	}

	fmt.Println("   Valid bookmark names:")
	for _, name := range validNames {
		testOpts := docxupdater.DefaultBookmarkOptions()
		testOpts.Position = docxupdater.PositionEnd
		if err := u.CreateBookmark(name, testOpts); err != nil {
			log.Printf("   ERROR: Unexpected failure for valid name '%s': %v", name, err)
		} else {
			fmt.Printf("   ✓ %s\n", name)
		}
	}

	fmt.Println("\n   Invalid bookmark names (should fail):")
	for _, name := range invalidNames {
		testOpts := docxupdater.DefaultBookmarkOptions()
		testOpts.Position = docxupdater.PositionEnd
		if err := u.CreateBookmark(name, testOpts); err != nil {
			fmt.Printf("   ✓ %s (correctly rejected: %v)\n", name, err)
		} else {
			log.Printf("   ERROR: Invalid name '%s' was incorrectly accepted", name)
		}
	}

	// Save the document
	outputPath := "output_bookmarks.docx"
	if err := u.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Printf("\n✓ Document saved successfully to: %s\n", outputPath)
	fmt.Println("\nKey Features Demonstrated:")
	fmt.Println("• Empty bookmarks as position markers")
	fmt.Println("• Bookmarks with text content")
	fmt.Println("• Multiple bookmarked sections")
	fmt.Println("• Table of contents with internal links")
	fmt.Println("• Wrapping existing text in bookmarks")
	fmt.Println("• Position-based bookmark insertion")
	fmt.Println("• Bookmark name validation")
	fmt.Println("\nOpen the document in Word and test the hyperlinks!")
}

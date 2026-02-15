package main

import (
	"fmt"
	"log"
	"time"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	fmt.Println("Testing paragraph insertion on docx_template.docx...")

	// Open the template
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	fmt.Println("✓ Opened template successfully")

	// Add document header
	if err := u.AddHeading(1, "Performance Report - Q1 2026", updater.PositionBeginning); err != nil {
		log.Fatalf("Failed to add main title: %v", err)
	}
	fmt.Println("✓ Added main title")

	// Add subtitle with date
	dateText := fmt.Sprintf("Generated on %s", time.Now().Format("January 2, 2006"))
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     dateText,
		Style:    updater.StyleSubtitle,
		Position: updater.PositionBeginning,
		Italic:   true,
	}); err != nil {
		log.Fatalf("Failed to add subtitle: %v", err)
	}
	fmt.Println("✓ Added subtitle with date")

	// Add executive summary section
	if err := u.AddHeading(2, "Executive Summary", updater.PositionEnd); err != nil {
		log.Fatalf("Failed to add section heading: %v", err)
	}
	fmt.Println("✓ Added Executive Summary heading")

	// Add summary content
	summaryParagraphs := []updater.ParagraphOptions{
		{
			Text:     "This report presents the comprehensive performance analysis for the first quarter of 2026. Our systems have demonstrated significant improvements across all key metrics.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:      "Key Highlights:",
			Style:     updater.StyleNormal,
			Position:  updater.PositionEnd,
			Bold:      true,
			Underline: true,
		},
		{
			Text:     "• System uptime improved to 99.97%",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "• Response times decreased by 28% compared to Q4 2025",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "• Error rate reduced to 0.08% (industry-leading)",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "• Customer satisfaction score: 4.8/5.0",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
	}

	if err := u.InsertParagraphs(summaryParagraphs); err != nil {
		log.Fatalf("Failed to add summary paragraphs: %v", err)
	}
	fmt.Println("✓ Added summary content with bullet points")

	// Add detailed analysis section
	if err := u.AddHeading(2, "Detailed Performance Analysis", updater.PositionEnd); err != nil {
		log.Fatalf("Failed to add analysis heading: %v", err)
	}
	fmt.Println("✓ Added Detailed Analysis heading")

	// Add subsections with formatted content
	detailedSections := []updater.ParagraphOptions{
		{
			Text:     "System Reliability",
			Style:    updater.StyleHeading3,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Our monitoring systems recorded 99.97% uptime across all services, exceeding our SLA commitments. The three brief incidents that occurred were resolved within 15 minutes, minimizing customer impact.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Performance Metrics",
			Style:    updater.StyleHeading3,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Average API response time: 127ms (target: <200ms)",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Italic:   true,
		},
		{
			Text:     "Peak concurrent users handled: 15,200 (previous record: 12,800)",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Italic:   true,
		},
		{
			Text:     "Data processing throughput: 2.4TB/day (18% increase)",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Italic:   true,
		},
		{
			Text:     "Security & Compliance",
			Style:    updater.StyleHeading3,
			Position: updater.PositionEnd,
		},
		{
			Text:     "All security audits passed with zero critical findings. We successfully completed SOC 2 Type II certification and implemented additional security measures including enhanced encryption and multi-factor authentication across all services.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
	}

	if err := u.InsertParagraphs(detailedSections); err != nil {
		log.Fatalf("Failed to add detailed sections: %v", err)
	}
	fmt.Println("✓ Added detailed analysis sections")

	// Add recommendations section
	if err := u.AddHeading(2, "Recommendations", updater.PositionEnd); err != nil {
		log.Fatalf("Failed to add recommendations heading: %v", err)
	}

	recommendations := []updater.ParagraphOptions{
		{
			Text:     "Based on the Q1 performance data, we recommend the following actions for Q2:",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Infrastructure Scaling:",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Bold:     true,
		},
		{
			Text:     "Increase database capacity by 30% to accommodate projected traffic growth. Current utilization is at 75% during peak hours.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Monitoring Enhancements:",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
			Bold:     true,
		},
		{
			Text:     "Deploy advanced anomaly detection to identify potential issues before they impact users. Implement predictive analytics for capacity planning.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
	}

	if err := u.InsertParagraphs(recommendations); err != nil {
		log.Fatalf("Failed to add recommendations: %v", err)
	}
	fmt.Println("✓ Added recommendations section")

	// Add conclusion with special formatting
	if err := u.AddHeading(2, "Conclusion", updater.PositionEnd); err != nil {
		log.Fatalf("Failed to add conclusion heading: %v", err)
	}

	conclusion := []updater.ParagraphOptions{
		{
			Text:     "Q1 2026 has been our strongest quarter to date, demonstrating the effectiveness of our infrastructure investments and optimization efforts. The team's dedication to excellence is reflected in every metric.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Looking ahead, we are well-positioned to handle increased demand while maintaining our high standards of reliability, performance, and security.",
			Style:    updater.StyleNormal,
			Position: updater.PositionEnd,
		},
		{
			Text:     "Note: This report is based on automated monitoring data collected from January 1 to March 31, 2026. All metrics have been independently verified and comply with industry standards.",
			Style:    updater.StyleQuote,
			Position: updater.PositionEnd,
			Italic:   true,
		},
	}

	if err := u.InsertParagraphs(conclusion); err != nil {
		log.Fatalf("Failed to add conclusion: %v", err)
	}
	fmt.Println("✓ Added conclusion")

	// Add footer note
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "---",
		Style:    updater.StyleNormal,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatalf("Failed to add separator: %v", err)
	}

	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "For questions or additional details, please contact the Performance Engineering team.",
		Style:    updater.StyleNormal,
		Position: updater.PositionEnd,
		Italic:   true,
	}); err != nil {
		log.Fatalf("Failed to add footer: %v", err)
	}
	fmt.Println("✓ Added footer")

	// Save the output
	outputPath := "outputs/paragraph_test_output.docx"
	if err := u.Save(outputPath); err != nil {
		log.Fatalf("Failed to save output: %v", err)
	}

	fmt.Println("\n✅ SUCCESS!")
	fmt.Printf("📄 Output saved to: %s\n", outputPath)
	fmt.Println("\nDocument includes:")
	fmt.Println("  • Title and subtitle with date")
	fmt.Println("  • Executive Summary with key highlights")
	fmt.Println("  • Detailed analysis with subsections")
	fmt.Println("  • Recommendations with bold headings")
	fmt.Println("  • Conclusion with quoted notes")
	fmt.Println("  • Multiple formatting styles (bold, italic, underline)")
	fmt.Println("  • Heading levels 1, 2, and 3")
}

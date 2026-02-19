package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// Fix empty XML elements in document properties that cause Word corruption
// Usage: go run fix_empty_properties.go <path-to-docx-file> [output-path]

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run fix_empty_properties.go <path-to-docx-file> [output-path]")
		fmt.Println("  Removes empty property elements that cause Word corruption issues")
		os.Exit(1)
	}

	inputPath := os.Args[1]
	outputPath := inputPath
	if len(os.Args) > 2 {
		outputPath = os.Args[2]
	} else {
		// Create backup and overwrite original
		backupPath := strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + "_backup" + filepath.Ext(inputPath)
		if err := copyFile(inputPath, backupPath); err != nil {
			fmt.Printf("❌ Failed to create backup: %v\n", err)
			os.Exit(1)
		}
		fmt.Printf("✓ Backup created: %s\n", backupPath)
	}

	fmt.Printf("Fixing: %s\n", inputPath)

	// Open the DOCX file (it's a ZIP archive)
	r, err := zip.OpenReader(inputPath)
	if err != nil {
		fmt.Printf("❌ Failed to open DOCX: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Create output buffer
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)

	fixedCount := 0

	// Process each file in the archive
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("❌ Failed to open %s: %v\n", f.Name, err)
			continue
		}

		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			fmt.Printf("❌ Failed to read %s: %v\n", f.Name, err)
			continue
		}

		// Fix empty elements in property files
		if strings.HasPrefix(f.Name, "docProps/") && strings.HasSuffix(f.Name, ".xml") {
			originalContent := string(content)
			fixedContent := removeEmptyElements(originalContent)
			if originalContent != fixedContent {
				fixedCount++
				fmt.Printf("  ✓ Fixed %s\n", f.Name)
				content = []byte(fixedContent)
			}
		}

		// Write file to new archive - preserve file header
		fh := &zip.FileHeader{
			Name:   f.Name,
			Method: f.Method,
		}
		fw, err := w.CreateHeader(fh)
		if err != nil {
			fmt.Printf("❌ Failed to write %s: %v\n", f.Name, err)
			continue
		}
		_, err = fw.Write(content)
		if err != nil {
			fmt.Printf("❌ Failed to write content for %s: %v\n", f.Name, err)
			continue
		}
	}

	if err := w.Close(); err != nil {
		fmt.Printf("❌ Failed to close archive: %v\n", err)
		os.Exit(1)
	}

	// Write the fixed document
	if err := os.WriteFile(outputPath, buf.Bytes(), 0644); err != nil {
		fmt.Printf("❌ Failed to write output: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("\n✓ Fixed %d property file(s)\n", fixedCount)
	fmt.Printf("✓ Output: %s\n", outputPath)
}

// removeEmptyElements removes empty XML elements like <tag></tag>
func removeEmptyElements(content string) string {
	// List of common property tags that should be removed if empty
	emptyTags := []string{
		"dc:title",
		"dc:subject",
		"dc:creator",
		"cp:keywords",
		"dc:description",
		"cp:lastModifiedBy",
		"cp:category",
		"dc:language",
		"Company",
		"Manager",
	}

	// Remove empty elements for each tag
	for _, tag := range emptyTags {
		// Pattern for completely empty element: <tag></tag>
		emptyPattern := fmt.Sprintf(`<%s></%s>`, regexp.QuoteMeta(tag), regexp.QuoteMeta(tag))
		content = regexp.MustCompile(emptyPattern).ReplaceAllString(content, "")

		// Pattern for empty element with whitespace: <tag> </tag> or <tag>   </tag>
		emptyWithSpacePattern := fmt.Sprintf(`<%s>\s*</%s>`, regexp.QuoteMeta(tag), regexp.QuoteMeta(tag))
		content = regexp.MustCompile(emptyWithSpacePattern).ReplaceAllString(content, "")
	}

	return content
}

// copyFile copies a file from src to dst
func copyFile(src, dst string) error {
	sourceFile, err := os.Open(src)
	if err != nil {
		return err
	}
	defer sourceFile.Close()

	destFile, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer destFile.Close()

	_, err = io.Copy(destFile, sourceFile)
	return err
}

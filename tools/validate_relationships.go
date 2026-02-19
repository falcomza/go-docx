package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
)

// Validates that all relationships in a DOCX point to existing files
// Usage: go run validate_relationships.go <path-to-docx-file>

type Relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Relationship []Relationship `xml:"Relationship"`
}

type Relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

type ContentTypes struct {
	XMLName  xml.Name   `xml:"Types"`
	Default  []Default  `xml:"Default"`
	Override []Override `xml:"Override"`
}

type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run validate_relationships.go <path-to-docx-file>")
		os.Exit(1)
	}

	docxPath := os.Args[1]
	fmt.Printf("Validating relationships in: %s\n\n", docxPath)

	// Open DOCX
	r, err := zip.OpenReader(docxPath)
	if err != nil {
		fmt.Printf("❌ Failed to open: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Build file map
	fileMap := make(map[string]bool)
	for _, f := range r.File {
		fileMap[f.Name] = true
	}

	errors := 0

	// Validate [Content_Types].xml
	fmt.Println("=== Validating Content Types ===")
	contentTypes, err := readContentTypes(r)
	if err != nil {
		fmt.Printf("❌ Failed to read [Content_Types].xml: %v\n", err)
		errors++
	} else {
		fmt.Printf("✓ Found %d default types and %d overrides\n", len(contentTypes.Default), len(contentTypes.Override))

		// Check that all Override PartNames exist
		for _, override := range contentTypes.Override {
			partName := strings.TrimPrefix(override.PartName, "/")
			if !fileMap[partName] {
				fmt.Printf("❌ Content type override references missing file: %s\n", override.PartName)
				errors++
			}
		}
	}

	// Find all _rels files and validate them
	fmt.Println("\n=== Validating Relationships ===")
	relsFiles := []string{}
	for name := range fileMap {
		if strings.HasSuffix(name, ".rels") {
			relsFiles = append(relsFiles, name)
		}
	}

	fmt.Printf("Found %d relationship files\n", len(relsFiles))

	for _, relsFile := range relsFiles {
		rels, err := readRelationships(r, relsFile)
		if err != nil {
			fmt.Printf("❌ Failed to read %s: %v\n", relsFile, err)
			errors++
			continue
		}

		// Determine the base directory for this rels file
		var baseDir string
		if relsFile == "_rels/.rels" {
			// Root rels file - targets are relative to root
			baseDir = ""
		} else if strings.Contains(relsFile, "_rels") {
			// Get the directory that contains the _rels folder
			baseDir = path.Dir(path.Dir(relsFile))
			if baseDir == "." {
				baseDir = ""
			}
		} else {
			baseDir = path.Dir(relsFile)
		}

		// Validate each relationship
		for _, rel := range rels.Relationship {
			target := rel.Target

			// Skip external relationships
			if strings.HasPrefix(target, "http://") || strings.HasPrefix(target, "https://") || strings.HasPrefix(target, "mailto:") {
				continue
			}

			// Resolve relative path
			var targetPath string
			if strings.HasPrefix(target, "/") {
				targetPath = strings.TrimPrefix(target, "/")
			} else {
				targetPath = path.Join(baseDir, target)
			}

			// Clean up path
			targetPath = path.Clean(targetPath)
			targetPath = strings.ReplaceAll(targetPath, "\\", "/")

			// Check if target exists
			if !fileMap[targetPath] {
				fmt.Printf("❌ %s: Relationship ID=%s targets missing file: %s (resolved to: %s)\n",
					relsFile, rel.ID, target, targetPath)
				errors++
			}
		}
	}

	// Summary
	fmt.Println("\n=== Summary ===")
	if errors == 0 {
		fmt.Println("✓ All relationships are valid")
		fmt.Println("✓ Document structure is consistent")
	} else {
		fmt.Printf("❌ Found %d relationship error(s)\n", errors)
		fmt.Println("\nThe document may be corrupted or have missing files.")
	}
}

func readContentTypes(r *zip.ReadCloser) (*ContentTypes, error) {
	var ct ContentTypes
	for _, f := range r.File {
		if f.Name == "[Content_Types].xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()

			data, err := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}

			if err := xml.Unmarshal(data, &ct); err != nil {
				return nil, err
			}
			return &ct, nil
		}
	}
	return nil, fmt.Errorf("Content_Types.xml not found")
}

func readRelationships(r *zip.ReadCloser, filename string) (*Relationships, error) {
	var rels Relationships
	for _, f := range r.File {
		if f.Name == filename {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()

			data, err := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}

			if err := xml.Unmarshal(data, &rels); err != nil {
				return nil, err
			}
			return &rels, nil
		}
	}
	return nil, fmt.Errorf("file not found")
}

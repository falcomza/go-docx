package docxchartupdater

import (
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// Updater updates chart caches and embedded workbook data inside a DOCX file.
type Updater struct {
	originalPath string
	tempDir      string
}

// New prepares a working copy of a DOCX for chart updates.
func New(docxPath string) (*Updater, error) {
	if docxPath == "" {
		return nil, errors.New("docx path is required")
	}
	if _, err := os.Stat(docxPath); err != nil {
		return nil, fmt.Errorf("stat docx: %w", err)
	}

	tempDir, err := os.MkdirTemp("", "docx-chart-updater-*")
	if err != nil {
		return nil, fmt.Errorf("create temp dir: %w", err)
	}

	if err := extractZip(docxPath, tempDir); err != nil {
		os.RemoveAll(tempDir)
		return nil, fmt.Errorf("extract docx: %w", err)
	}

	return &Updater{originalPath: docxPath, tempDir: tempDir}, nil
}

// Cleanup removes temporary workspace.
func (u *Updater) Cleanup() error {
	if u == nil || u.tempDir == "" {
		return nil
	}
	return os.RemoveAll(u.tempDir)
}

// UpdateChart updates one chart by index (1-based).
func (u *Updater) UpdateChart(chartIndex int, data ChartData) error {
	if u == nil {
		return errors.New("updater is nil")
	}
	if chartIndex < 1 {
		return errors.New("chart index must be >= 1")
	}
	if err := validateChartData(data); err != nil {
		return err
	}

	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	if _, err := os.Stat(chartPath); err != nil {
		return fmt.Errorf("chart file does not exist: %w", err)
	}

	if err := updateChartXML(chartPath, data); err != nil {
		return fmt.Errorf("update chart xml: %w", err)
	}

	xlsxPath, err := u.findWorkbookPathForChart(chartIndex)
	if err != nil {
		return fmt.Errorf("resolve embedded workbook: %w", err)
	}
	if err := updateEmbeddedWorkbook(xlsxPath, data); err != nil {
		return fmt.Errorf("update embedded workbook: %w", err)
	}

	return nil
}

// Save writes the updated DOCX to outputPath.
func (u *Updater) Save(outputPath string) error {
	if u == nil {
		return errors.New("updater is nil")
	}
	if outputPath == "" {
		return errors.New("output path is required")
	}
	if err := os.MkdirAll(filepath.Dir(outputPath), 0o755); err != nil {
		return fmt.Errorf("create output dir: %w", err)
	}
	if err := createZipFromDir(u.tempDir, outputPath); err != nil {
		return fmt.Errorf("create output docx: %w", err)
	}
	return nil
}

func validateChartData(data ChartData) error {
	if len(data.Categories) == 0 {
		return errors.New("categories cannot be empty")
	}
	if len(data.Series) == 0 {
		return errors.New("series cannot be empty")
	}
	for i, s := range data.Series {
		if strings.TrimSpace(s.Name) == "" {
			return fmt.Errorf("series[%d] name cannot be empty", i)
		}
		if len(s.Values) != len(data.Categories) {
			return fmt.Errorf("series[%d] values length (%d) must match categories length (%d)", i, len(s.Values), len(data.Categories))
		}
	}
	return nil
}

func (u *Updater) findWorkbookPathForChart(chartIndex int) (string, error) {
	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	rawChart, err := os.ReadFile(chartPath)
	if err != nil {
		return "", fmt.Errorf("read chart xml: %w", err)
	}

	relID := externalDataRelID(rawChart)
	if relID != "" {
		relsPath := filepath.Join(u.tempDir, "word", "charts", "_rels", fmt.Sprintf("chart%d.xml.rels", chartIndex))
		target, err := findRelationshipTarget(relsPath, relID)
		if err == nil && target != "" {
			// Relationship targets are relative to the source part (chart#.xml), not the .rels folder.
			resolved := filepath.Clean(filepath.Join(filepath.Dir(chartPath), filepath.FromSlash(target)))
			if _, statErr := os.Stat(resolved); statErr == nil {
				return resolved, nil
			}
		}
	}

	fallback := filepath.Join(u.tempDir, "word", "embeddings")
	entries, err := os.ReadDir(fallback)
	if err != nil {
		return "", fmt.Errorf("read fallback embeddings dir: %w", err)
	}
	var candidates []string
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := strings.ToLower(entry.Name())
		if strings.HasSuffix(name, ".xlsx") {
			candidates = append(candidates, filepath.Join(fallback, entry.Name()))
		}
	}
	if len(candidates) == 0 {
		return "", errors.New("no embedded xlsx file found")
	}
	sort.Strings(candidates)
	return candidates[0], nil
}

func externalDataRelID(chartXML []byte) string {
	content := string(chartXML)
	const marker = "<c:externalData"
	start := strings.Index(content, marker)
	if start == -1 {
		return ""
	}
	end := strings.Index(content[start:], ">")
	if end == -1 {
		return ""
	}
	tag := content[start : start+end]
	const relAttr = `r:id="`
	attrIdx := strings.Index(tag, relAttr)
	if attrIdx == -1 {
		return ""
	}
	valueStart := attrIdx + len(relAttr)
	valueEnd := strings.Index(tag[valueStart:], `"`)
	if valueEnd == -1 {
		return ""
	}
	return tag[valueStart : valueStart+valueEnd]
}

type relationships struct {
	XMLName       xml.Name       `xml:"Relationships"`
	Relationships []relationship `xml:"Relationship"`
}

type relationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
}

func findRelationshipTarget(relsPath, relationshipID string) (string, error) {
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read relationships: %w", err)
	}
	var rels relationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return "", fmt.Errorf("parse relationships: %w", err)
	}
	for _, rel := range rels.Relationships {
		if rel.ID == relationshipID {
			return rel.Target, nil
		}
	}
	return "", fmt.Errorf("relationship %s not found", relationshipID)
}

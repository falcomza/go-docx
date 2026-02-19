package docxupdater

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"
)

// CoreProperties represents core document properties
type CoreProperties struct {
	// Title of the document
	Title string

	// Subject of the document
	Subject string

	// Creator/Author of the document
	Creator string

	// Keywords for the document (comma-separated or slice)
	Keywords string

	// Description/Comments
	Description string

	// Category of the document
	Category string

	// Created date (if empty, uses current time)
	Created time.Time

	// Modified date (if empty, uses current time)
	Modified time.Time

	// LastModifiedBy user name
	LastModifiedBy string

	// Revision number (version)
	Revision string
}

// AppProperties represents application-specific document properties
type AppProperties struct {
	// Company name
	Company string

	// Manager name
	Manager string

	// Application name (typically Microsoft Word)
	Application string

	// AppVersion (e.g., "16.0000")
	AppVersion string
}

// CustomProperty represents a custom document property
type CustomProperty struct {
	// Name of the property
	Name string

	// Value of the property (string, int, float64, bool, or time.Time)
	Value any

	// Type is inferred from Value, but can be explicitly set
	// Valid types: "lpwstr" (string), "i4" (int), "r8" (float), "bool", "date"
	Type string
}

// SetCoreProperties sets the core document properties
func (u *Updater) SetCoreProperties(props CoreProperties) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	corePath := filepath.Join(u.tempDir, "docProps", "core.xml")

	// Read existing core.xml or create new
	var content string
	if raw, err := os.ReadFile(corePath); err == nil {
		content = string(raw)
	} else {
		content = u.generateDefaultCoreXML()
	}

	// Update properties
	content = u.updateCoreProperty(content, "dc:title", props.Title)
	content = u.updateCoreProperty(content, "dc:subject", props.Subject)
	content = u.updateCoreProperty(content, "dc:creator", props.Creator)
	content = u.updateCoreProperty(content, "cp:keywords", props.Keywords)
	content = u.updateCoreProperty(content, "dc:description", props.Description)
	content = u.updateCoreProperty(content, "cp:category", props.Category)
	content = u.updateCoreProperty(content, "cp:lastModifiedBy", props.LastModifiedBy)
	content = u.updateCoreProperty(content, "cp:revision", props.Revision)

	// Update dates with proper attributes
	if !props.Created.IsZero() {
		content = u.updateCoreDateProperty(content, "dcterms:created", props.Created.Format(time.RFC3339))
	}
	if !props.Modified.IsZero() {
		content = u.updateCoreDateProperty(content, "dcterms:modified", props.Modified.Format(time.RFC3339))
	} else {
		// Always update modified to current time
		content = u.updateCoreDateProperty(content, "dcterms:modified", time.Now().Format(time.RFC3339))
	}

	// Write updated core.xml
	if err := os.WriteFile(corePath, []byte(content), 0o644); err != nil {
		return &DocxError{
			Code:    "PROPERTIES_ERROR",
			Message: "failed to write core properties",
			Err:     err,
		}
	}

	return nil
}

// SetAppProperties sets the application-specific document properties
func (u *Updater) SetAppProperties(props AppProperties) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	appPath := filepath.Join(u.tempDir, "docProps", "app.xml")

	// Read existing app.xml or create new
	var content string
	if raw, err := os.ReadFile(appPath); err == nil {
		content = string(raw)
	} else {
		content = u.generateDefaultAppXML()
	}

	// Update properties
	if props.Company != "" {
		content = u.updateAppProperty(content, "Company", props.Company)
	}
	if props.Manager != "" {
		content = u.updateAppProperty(content, "Manager", props.Manager)
	}
	if props.Application != "" {
		content = u.updateAppProperty(content, "Application", props.Application)
	}
	if props.AppVersion != "" {
		content = u.updateAppProperty(content, "AppVersion", props.AppVersion)
	}

	// Write updated app.xml
	if err := os.WriteFile(appPath, []byte(content), 0o644); err != nil {
		return &DocxError{
			Code:    "PROPERTIES_ERROR",
			Message: "failed to write app properties",
			Err:     err,
		}
	}

	return nil
}

// SetCustomProperties sets custom document properties
func (u *Updater) SetCustomProperties(properties []CustomProperty) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	customPath := filepath.Join(u.tempDir, "docProps", "custom.xml")

	// Generate custom.xml content
	content := u.generateCustomPropertiesXML(properties)

	// Ensure docProps directory exists
	if err := os.MkdirAll(filepath.Dir(customPath), 0o755); err != nil {
		return &DocxError{
			Code:    "PROPERTIES_ERROR",
			Message: "failed to create docProps directory",
			Err:     err,
		}
	}

	// Write custom.xml
	if err := os.WriteFile(customPath, []byte(content), 0o644); err != nil {
		return &DocxError{
			Code:    "PROPERTIES_ERROR",
			Message: "failed to write custom properties",
			Err:     err,
		}
	}

	// Add custom.xml to content types if not present
	if err := u.addCustomPropertiesContentType(); err != nil {
		return err
	}

	// Add custom.xml relationship if not present
	if err := u.addCustomPropertiesRelationship(); err != nil {
		return err
	}

	return nil
}

// GetCoreProperties retrieves core document properties
func (u *Updater) GetCoreProperties() (*CoreProperties, error) {
	if u == nil {
		return nil, fmt.Errorf("updater is nil")
	}

	corePath := filepath.Join(u.tempDir, "docProps", "core.xml")
	raw, err := os.ReadFile(corePath)
	if err != nil {
		return nil, &DocxError{
			Code:    "PROPERTIES_ERROR",
			Message: "failed to read core properties",
			Err:     err,
		}
	}

	content := string(raw)
	props := &CoreProperties{}

	props.Title = u.extractCoreProperty(content, "dc:title")
	props.Subject = u.extractCoreProperty(content, "dc:subject")
	props.Creator = u.extractCoreProperty(content, "dc:creator")
	props.Keywords = u.extractCoreProperty(content, "cp:keywords")
	props.Description = u.extractCoreProperty(content, "dc:description")
	props.Category = u.extractCoreProperty(content, "cp:category")
	props.LastModifiedBy = u.extractCoreProperty(content, "cp:lastModifiedBy")
	props.Revision = u.extractCoreProperty(content, "cp:revision")

	// Parse dates
	if created := u.extractCoreProperty(content, "dcterms:created"); created != "" {
		if t, err := time.Parse(time.RFC3339, created); err == nil {
			props.Created = t
		}
	}
	if modified := u.extractCoreProperty(content, "dcterms:modified"); modified != "" {
		if t, err := time.Parse(time.RFC3339, modified); err == nil {
			props.Modified = t
		}
	}

	return props, nil
}

// updateCoreProperty updates or adds a core property in the XML
func (u *Updater) updateCoreProperty(content, property, value string) string {
	// Check if property exists
	pattern := fmt.Sprintf(`<%s[^>]*>.*?</%s>`, regexp.QuoteMeta(property), regexp.QuoteMeta(property))
	re := regexp.MustCompile(pattern)

	if value == "" {
		// If value is empty, remove the property if it exists
		if re.MatchString(content) {
			content = re.ReplaceAllString(content, "")
		}
		return content
	}

	escapedValue := escapeXML(value)

	if re.MatchString(content) {
		// Update existing
		replacement := fmt.Sprintf(`<%s>%s</%s>`, property, escapedValue, property)
		content = re.ReplaceAllString(content, replacement)
	} else {
		// Add new property before </cp:coreProperties>
		newProp := fmt.Sprintf(`<%s>%s</%s>`, property, escapedValue, property)
		content = strings.Replace(content, "</cp:coreProperties>", newProp+"</cp:coreProperties>", 1)
	}

	return content
}

// updateCoreDateProperty updates or adds a date property in core.xml with proper attributes
func (u *Updater) updateCoreDateProperty(content, property, value string) string {
	if value == "" {
		return content
	}

	// Both dcterms:created and dcterms:modified need xsi:type attribute
	var element string
	if property == "dcterms:created" || property == "dcterms:modified" {
		element = fmt.Sprintf(`<%s xsi:type="dcterms:W3CDTF">%s</%s>`, property, value, property)
	} else {
		element = fmt.Sprintf(`<%s>%s</%s>`, property, value, property)
	}

	// Check if property exists (with or without attributes)
	pattern := fmt.Sprintf(`<%s[^>]*>.*?</%s>`, regexp.QuoteMeta(property), regexp.QuoteMeta(property))
	re := regexp.MustCompile(pattern)

	if re.MatchString(content) {
		// Update existing
		content = re.ReplaceAllString(content, element)
	} else {
		// Add new property before </cp:coreProperties>
		content = strings.Replace(content, "</cp:coreProperties>", element+"</cp:coreProperties>", 1)
	}

	return content
}

// updateAppProperty updates or adds an app property in the XML
func (u *Updater) updateAppProperty(content, property, value string) string {
	// Check if property exists
	pattern := fmt.Sprintf(`<%s[^>]*>.*?</%s>`, regexp.QuoteMeta(property), regexp.QuoteMeta(property))
	re := regexp.MustCompile(pattern)

	if value == "" {
		// If value is empty, remove the property if it exists
		if re.MatchString(content) {
			content = re.ReplaceAllString(content, "")
		}
		return content
	}

	escapedValue := escapeXML(value)

	if re.MatchString(content) {
		// Update existing
		replacement := fmt.Sprintf(`<%s>%s</%s>`, property, escapedValue, property)
		content = re.ReplaceAllString(content, replacement)
	} else {
		// Add new property before </Properties>
		newProp := fmt.Sprintf(`<%s>%s</%s>`, property, escapedValue, property)
		content = strings.Replace(content, "</Properties>", newProp+"</Properties>", 1)
	}

	return content
}

// extractCoreProperty extracts a property value from core.xml
func (u *Updater) extractCoreProperty(content, property string) string {
	pattern := fmt.Sprintf(`<%s[^>]*>(.*?)</%s>`, regexp.QuoteMeta(property), regexp.QuoteMeta(property))
	re := regexp.MustCompile(pattern)
	if match := re.FindStringSubmatch(content); len(match) > 1 {
		return match[1]
	}
	return ""
}

// generateDefaultCoreXML creates a minimal core.xml
func (u *Updater) generateDefaultCoreXML() string {
	now := time.Now().Format(time.RFC3339)
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<cp:revision>1</cp:revision>
<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>`, now, now)
}

// generateDefaultAppXML creates a minimal app.xml
func (u *Updater) generateDefaultAppXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Word</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0000</AppVersion>
</Properties>`
}

// generateCustomPropertiesXML creates custom.xml content
func (u *Updater) generateCustomPropertiesXML(properties []CustomProperty) string {
	var buf strings.Builder

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`)
	buf.WriteString("\n")

	for i, prop := range properties {
		pid := i + 2 // PIDs start at 2
		buf.WriteString(fmt.Sprintf(`<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="%s">`, pid, escapeXML(prop.Name)))

		// Determine type and format value
		propType := prop.Type
		if propType == "" {
			propType = u.inferCustomPropertyType(prop.Value)
		}

		switch propType {
		case "lpwstr": // String
			buf.WriteString(fmt.Sprintf(`<vt:lpwstr>%s</vt:lpwstr>`, escapeXML(fmt.Sprintf("%v", prop.Value))))
		case "i4": // Integer
			buf.WriteString(fmt.Sprintf(`<vt:i4>%v</vt:i4>`, prop.Value))
		case "r8": // Float
			buf.WriteString(fmt.Sprintf(`<vt:r8>%v</vt:r8>`, prop.Value))
		case "bool": // Boolean
			buf.WriteString(fmt.Sprintf(`<vt:bool>%v</vt:bool>`, prop.Value))
		case "date": // Date
			if t, ok := prop.Value.(time.Time); ok {
				// Convert to Windows FILETIME format (100-nanosecond intervals since Jan 1, 1601)
				filetime := timeToFiletime(t)
				buf.WriteString(fmt.Sprintf(`<vt:filetime>%s</vt:filetime>`, filetime))
			} else {
				buf.WriteString(fmt.Sprintf(`<vt:lpwstr>%v</vt:lpwstr>`, prop.Value))
			}
		default:
			buf.WriteString(fmt.Sprintf(`<vt:lpwstr>%v</vt:lpwstr>`, prop.Value))
		}

		buf.WriteString(`</property>`)
		buf.WriteString("\n")
	}

	buf.WriteString(`</Properties>`)

	return buf.String()
}

// inferCustomPropertyType infers the vt type from Go value
func (u *Updater) inferCustomPropertyType(value any) string {
	switch value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		return "i4"
	case float32, float64:
		return "r8"
	case bool:
		return "bool"
	case time.Time:
		return "date"
	default:
		return "lpwstr" // String is default
	}
}

// addCustomPropertiesContentType adds custom.xml to [Content_Types].xml
func (u *Updater) addCustomPropertiesContentType() error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")

	raw, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read content types: %w", err)
	}

	content := string(raw)

	// Check if already exists
	if strings.Contains(content, "docProps/custom.xml") {
		return nil
	}

	override := `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`

	// Insert before closing </Types>
	content = strings.Replace(content, "</Types>", override+"</Types>", 1)

	// Write updated content types
	if err := os.WriteFile(contentTypesPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write content types: %w", err)
	}

	return nil
}

// addCustomPropertiesRelationship adds custom.xml relationship to _rels/.rels
func (u *Updater) addCustomPropertiesRelationship() error {
	relsPath := filepath.Join(u.tempDir, "_rels", ".rels")

	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read relationships: %w", err)
	}

	content := string(raw)

	// Check if relationship already exists
	if strings.Contains(content, "docProps/custom.xml") {
		return nil
	}

	// Find next available relationship ID
	nextID := u.getNextRelationshipID(content)
	relID := fmt.Sprintf("rId%d", nextID)

	newRel := fmt.Sprintf(
		`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>`,
		relID,
	)

	// Insert before closing </Relationships>
	content = strings.Replace(content, "</Relationships>", newRel+"</Relationships>", 1)

	// Write updated relationships
	if err := os.WriteFile(relsPath, []byte(content), 0o644); err != nil {
		return fmt.Errorf("write relationships: %w", err)
	}

	return nil
}

// timeToFiletime converts time.Time to Windows FILETIME format
// FILETIME is the number of 100-nanosecond intervals since January 1, 1601 UTC
func timeToFiletime(t time.Time) string {
	// Office Open XML uses ISO 8601 format for vt:filetime
	return t.UTC().Format(time.RFC3339)
}

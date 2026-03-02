package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// CommentOptions defines options for inserting a comment
type CommentOptions struct {
	// Text is the comment content
	Text string

	// Author is the comment author name
	Author string

	// Initials is the author's initials (derived from Author if empty)
	Initials string

	// Anchor is the text in the document to attach the comment to.
	// The comment range will span the paragraph containing this text.
	Anchor string
}

// Comment represents an existing comment in the document
type Comment struct {
	ID       int
	Author   string
	Initials string
	Date     string
	Text     string
}

var (
	commentIDPattern     = regexp.MustCompile(`<w:comment[^>]*w:id="(\d+)"`)
	commentAuthorPattern = regexp.MustCompile(`w:author="([^"]*)"`)
	commentInitPattern   = regexp.MustCompile(`w:initials="([^"]*)"`)
	commentDatePattern   = regexp.MustCompile(`w:date="([^"]*)"`)
	commentBlockPattern  = regexp.MustCompile(`(?s)<w:comment\s+[^>]*>.*?</w:comment>`)
	commentTextPattern   = regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
)

// InsertComment adds a comment to the document.
// The comment range spans the paragraph containing the anchor text.
func (u *Updater) InsertComment(opts CommentOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Text == "" {
		return fmt.Errorf("comment text cannot be empty")
	}
	if opts.Anchor == "" {
		return fmt.Errorf("anchor text cannot be empty")
	}
	if opts.Author == "" {
		opts.Author = "Author"
	}
	if opts.Initials == "" && len(opts.Author) > 0 {
		opts.Initials = string(opts.Author[0])
	}

	commentID, err := u.ensureCommentsXML()
	if err != nil {
		return fmt.Errorf("ensure comments.xml: %w", err)
	}

	if err := u.addCommentContent(commentID, opts); err != nil {
		return fmt.Errorf("add comment content: %w", err)
	}

	if err := u.insertCommentMarkers(opts.Anchor, commentID); err != nil {
		return fmt.Errorf("insert comment markers: %w", err)
	}

	return nil
}

// GetComments reads all comments from the document.
// Returns nil if the document has no comments.
func (u *Updater) GetComments() ([]Comment, error) {
	if u == nil {
		return nil, fmt.Errorf("updater is nil")
	}

	commentsPath := filepath.Join(u.tempDir, "word", "comments.xml")
	raw, err := os.ReadFile(commentsPath)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, fmt.Errorf("read comments.xml: %w", err)
	}

	return parseComments(raw), nil
}

// ensureCommentsXML creates comments.xml if it doesn't exist and returns the next available ID.
func (u *Updater) ensureCommentsXML() (int, error) {
	commentsPath := filepath.Join(u.tempDir, "word", "comments.xml")

	if _, err := os.Stat(commentsPath); os.IsNotExist(err) {
		content := generateInitialCommentsXML()
		if err := atomicWriteFile(commentsPath, content, 0o644); err != nil {
			return 0, fmt.Errorf("write comments.xml: %w", err)
		}

		if err := u.addNoteRelationship("comments.xml", "comments"); err != nil {
			return 0, fmt.Errorf("add comments relationship: %w", err)
		}

		if err := u.addNoteContentType("comments.xml", "comments"); err != nil {
			return 0, fmt.Errorf("add comments content type: %w", err)
		}

		return 1, nil
	}

	raw, err := os.ReadFile(commentsPath)
	if err != nil {
		return 0, fmt.Errorf("read comments.xml: %w", err)
	}

	return getNextCommentID(raw), nil
}

// generateInitialCommentsXML creates a new empty comments.xml
func generateInitialCommentsXML() []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `)
	buf.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString("\n")
	buf.WriteString(`</w:comments>`)

	return buf.Bytes()
}

// addCommentContent adds a comment entry to comments.xml
func (u *Updater) addCommentContent(id int, opts CommentOptions) error {
	commentsPath := filepath.Join(u.tempDir, "word", "comments.xml")
	raw, err := os.ReadFile(commentsPath)
	if err != nil {
		return fmt.Errorf("read comments.xml: %w", err)
	}

	commentXML := generateCommentEntry(id, opts)

	closeTag := []byte("</w:comments>")
	closeIdx := bytes.LastIndex(raw, closeTag)
	if closeIdx == -1 {
		return fmt.Errorf("could not find </w:comments> tag")
	}

	result := make([]byte, 0, len(raw)+len(commentXML)+1)
	result = append(result, raw[:closeIdx]...)
	result = append(result, commentXML...)
	result = append(result, '\n')
	result = append(result, raw[closeIdx:]...)

	return atomicWriteFile(commentsPath, result, 0o644)
}

// generateCommentEntry creates the XML for a single comment
func generateCommentEntry(id int, opts CommentOptions) []byte {
	var buf bytes.Buffer

	dateStr := time.Now().UTC().Format(time.RFC3339)

	buf.WriteString(fmt.Sprintf(
		`<w:comment w:id="%d" w:author="%s" w:date="%s" w:initials="%s">`,
		id, xmlEscape(opts.Author), dateStr, xmlEscape(opts.Initials)))
	buf.WriteString("<w:p>")
	buf.WriteString(`<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>`)

	buf.WriteString("<w:r>")
	buf.WriteString(`<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>`)
	buf.WriteString("<w:annotationRef/>")
	buf.WriteString("</w:r>")

	buf.WriteString("<w:r>")
	buf.WriteString(fmt.Sprintf(`<w:t xml:space="preserve"> %s</w:t>`, xmlEscape(opts.Text)))
	buf.WriteString("</w:r>")

	buf.WriteString("</w:p>")
	buf.WriteString("</w:comment>")

	return buf.Bytes()
}

// insertCommentMarkers inserts commentRangeStart, commentRangeEnd, and commentReference
// into document.xml around the paragraph containing the anchor text.
func (u *Updater) insertCommentMarkers(anchor string, commentID int) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	paraStart, paraEnd, err := findParagraphRangeByAnchor(raw, anchor)
	if err != nil {
		return fmt.Errorf("find anchor: %w", err)
	}

	pContent := string(raw[paraStart:paraEnd])

	// Insert commentRangeStart after <w:pPr>...</w:pPr> if present,
	// otherwise after the opening <w:p> tag
	var insertStartOffset int
	if pprEnd := strings.Index(pContent, "</w:pPr>"); pprEnd >= 0 {
		insertStartOffset = pprEnd + len("</w:pPr>")
	} else {
		pOpenEnd := strings.Index(pContent, ">")
		if pOpenEnd < 0 {
			return fmt.Errorf("invalid paragraph XML")
		}
		insertStartOffset = pOpenEnd + 1
	}
	insertStartPos := paraStart + insertStartOffset
	insertEndPos := paraEnd - len("</w:p>")

	rangeStartXML := fmt.Sprintf(`<w:commentRangeStart w:id="%d"/>`, commentID)
	rangeEndXML := fmt.Sprintf(
		`<w:commentRangeEnd w:id="%d"/>`+
			`<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>`+
			`<w:commentReference w:id="%d"/></w:r>`,
		commentID, commentID)

	result := make([]byte, 0, len(raw)+len(rangeStartXML)+len(rangeEndXML))
	result = append(result, raw[:insertStartPos]...)
	result = append(result, []byte(rangeStartXML)...)
	result = append(result, raw[insertStartPos:insertEndPos]...)
	result = append(result, []byte(rangeEndXML)...)
	result = append(result, raw[insertEndPos:]...)

	return atomicWriteFile(docPath, result, 0o644)
}

// getNextCommentID finds the next available comment ID in comments.xml
func getNextCommentID(raw []byte) int {
	matches := commentIDPattern.FindAllSubmatch(raw, -1)

	maxID := 0
	for _, match := range matches {
		if len(match) > 1 {
			id, err := strconv.Atoi(string(match[1]))
			if err != nil {
				continue
			}
			if id > maxID {
				maxID = id
			}
		}
	}

	return maxID + 1
}

// parseComments extracts all comments from comments.xml content
func parseComments(raw []byte) []Comment {
	var comments []Comment
	content := string(raw)

	blocks := commentBlockPattern.FindAllString(content, -1)

	for _, block := range blocks {
		c := Comment{}

		if m := commentIDPattern.FindStringSubmatch(block); len(m) > 1 {
			c.ID, _ = strconv.Atoi(m[1])
		}
		if m := commentAuthorPattern.FindStringSubmatch(block); len(m) > 1 {
			c.Author = m[1]
		}
		if m := commentInitPattern.FindStringSubmatch(block); len(m) > 1 {
			c.Initials = m[1]
		}
		if m := commentDatePattern.FindStringSubmatch(block); len(m) > 1 {
			c.Date = m[1]
		}

		var texts []string
		textMatches := commentTextPattern.FindAllStringSubmatch(block, -1)
		for _, tm := range textMatches {
			if len(tm) > 1 {
				texts = append(texts, tm[1])
			}
		}
		c.Text = strings.Join(texts, "")

		if c.ID > 0 {
			comments = append(comments, c)
		}
	}

	return comments
}

package godocx

import (
	"archive/zip"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"strings"
)

// maxExtractedFileSize is the largest a single decompressed zip entry may be.
// 256 MiB covers any realistic DOCX part while preventing zip-bomb exhaustion.
const maxExtractedFileSize = 256 << 20 // 256 MiB

func extractZip(zipPath, destDir string) error {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return fmt.Errorf("open zip: %w", err)
	}
	defer r.Close()

	cleanDest := filepath.Clean(destDir)

	for _, f := range r.File {
		target := filepath.Join(destDir, filepath.FromSlash(f.Name))

		// Zip Slip protection: ensure the target path stays within destDir
		cleanTarget := filepath.Clean(target)
		if !strings.HasPrefix(cleanTarget+string(os.PathSeparator), cleanDest+string(os.PathSeparator)) &&
			cleanTarget != cleanDest {
			return fmt.Errorf("zip entry %s escapes target directory", f.Name)
		}

		if f.FileInfo().IsDir() {
			if err := os.MkdirAll(target, 0755); err != nil {
				return fmt.Errorf("create dir %s: %w", target, err)
			}
			continue
		}

		if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
			return fmt.Errorf("create parent dir for %s: %w", target, err)
		}

		rc, err := f.Open()
		if err != nil {
			return fmt.Errorf("open zip entry %s: %w", f.Name, err)
		}

		out, err := os.Create(target)
		if err != nil {
			rc.Close()
			return fmt.Errorf("create file %s: %w", target, err)
		}

		// Limit decompressed size to guard against zip-bomb payloads.
		lr := io.LimitReader(rc, maxExtractedFileSize+1)
		n, copyErr := io.Copy(out, lr)
		if copyErr != nil {
			out.Close()
			rc.Close()
			return fmt.Errorf("copy zip entry %s: %w", f.Name, copyErr)
		}
		if n > maxExtractedFileSize {
			out.Close()
			rc.Close()
			return fmt.Errorf("zip entry %s exceeds maximum allowed size (%d bytes)", f.Name, maxExtractedFileSize)
		}

		if err := out.Close(); err != nil {
			rc.Close()
			return fmt.Errorf("close file %s: %w", target, err)
		}

		if err := rc.Close(); err != nil {
			return fmt.Errorf("close zip entry %s: %w", f.Name, err)
		}
	}

	return nil
}

func createZipFromDir(sourceDir, outZipPath string) error {
	out, err := os.Create(outZipPath)
	if err != nil {
		return fmt.Errorf("create output zip: %w", err)
	}

	if err := writeZipFromDir(sourceDir, out); err != nil {
		out.Close()
		return err
	}

	if err := out.Close(); err != nil {
		return fmt.Errorf("close output zip: %w", err)
	}

	return nil
}

// writeZipFromDir writes a zip archive of sourceDir to the given writer.
func writeZipFromDir(sourceDir string, w io.Writer) error {
	zw := zip.NewWriter(w)

	walkErr := filepath.WalkDir(sourceDir, func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if path == sourceDir || d.IsDir() {
			return nil
		}

		rel, err := filepath.Rel(sourceDir, path)
		if err != nil {
			return fmt.Errorf("relative path for %s: %w", path, err)
		}

		zipPath := filepath.ToSlash(rel)
		ew, err := zw.Create(zipPath)
		if err != nil {
			return fmt.Errorf("create zip entry %s: %w", zipPath, err)
		}

		f, err := os.Open(path)
		if err != nil {
			return fmt.Errorf("open source file %s: %w", path, err)
		}

		if _, err := io.Copy(ew, f); err != nil {
			f.Close()
			return fmt.Errorf("write zip entry %s: %w", zipPath, err)
		}

		if err := f.Close(); err != nil {
			return fmt.Errorf("close source file %s: %w", path, err)
		}

		return nil
	})
	if walkErr != nil {
		zw.Close()
		return walkErr
	}

	if err := zw.Close(); err != nil {
		return fmt.Errorf("close zip writer: %w", err)
	}

	return nil
}

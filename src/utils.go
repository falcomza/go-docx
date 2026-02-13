package docxchartupdater

import (
	"archive/zip"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
)

func extractZip(zipPath, destDir string) error {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return fmt.Errorf("open zip: %w", err)
	}
	defer r.Close()

	for _, f := range r.File {
		target := filepath.Join(destDir, filepath.FromSlash(f.Name))

		if f.FileInfo().IsDir() {
			if err := os.MkdirAll(target, 0o755); err != nil {
				return fmt.Errorf("create dir %s: %w", target, err)
			}
			continue
		}

		if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
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

		_, copyErr := io.Copy(out, rc)
		closeErr := out.Close()
		rcCloseErr := rc.Close()
		if copyErr != nil {
			return fmt.Errorf("copy zip entry %s: %w", f.Name, copyErr)
		}
		if closeErr != nil {
			return fmt.Errorf("close file %s: %w", target, closeErr)
		}
		if rcCloseErr != nil {
			return fmt.Errorf("close zip entry %s: %w", f.Name, rcCloseErr)
		}
	}

	return nil
}

func createZipFromDir(sourceDir, outZipPath string) error {
	out, err := os.Create(outZipPath)
	if err != nil {
		return fmt.Errorf("create output zip: %w", err)
	}
	defer out.Close()

	zw := zip.NewWriter(out)
	defer zw.Close()

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
		w, err := zw.Create(zipPath)
		if err != nil {
			return fmt.Errorf("create zip entry %s: %w", zipPath, err)
		}

		f, err := os.Open(path)
		if err != nil {
			return fmt.Errorf("open source file %s: %w", path, err)
		}
		defer f.Close()

		if _, err := io.Copy(w, f); err != nil {
			return fmt.Errorf("write zip entry %s: %w", zipPath, err)
		}

		return nil
	})
	if walkErr != nil {
		return walkErr
	}

	if err := zw.Close(); err != nil {
		return fmt.Errorf("close zip writer: %w", err)
	}

	if err := out.Close(); err != nil {
		return fmt.Errorf("close output zip: %w", err)
	}

	return nil
}

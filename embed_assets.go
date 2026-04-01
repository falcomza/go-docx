package godocx

import _ "embed"

// defaultExcelIconPNG is the built-in Excel icon used when no custom icon is provided.
// To replace it: swap out assets/excel_icon.png and rebuild the package.
//
//go:embed assets/excel_icon.png
var defaultExcelIconPNG []byte

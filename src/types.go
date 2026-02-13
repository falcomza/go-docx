package docxchartupdater

// ChartData defines chart categories and series values.
type ChartData struct {
	Categories []string
	Series     []SeriesData
}

// SeriesData defines one chart series.
type SeriesData struct {
	Name   string
	Values []float64
}

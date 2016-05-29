package kwdataexpoter

import "testing"

const filename string = "sample.xlsx"

func TestExportFile(t *testing.T) {
	ExportFile(filename)
}

func BenchmarkExportFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		ExportFile(filename)
	}
}

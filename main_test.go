package main

import "github.com/tealeg/xlsx"
import "testing"

const filename string = "sample.xlsx"

func TestTrimString(t *testing.T) {
	var sheet xlsx.Sheet
	row := sheet.AddRow()
	cell := row.AddCell()

	cell.SetString("#JSON ")
	if result := TrimString(cell); result != "JSON" {
		t.Error("TrimString '#JSON ' should be 'JSON' but", result)
	}
}

func TestIsComment(t *testing.T) {
	var sheet xlsx.Sheet
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.SetString("#This is comment")
	if result := IsComment(cell); result != true {
		t.Error("'#This is comment' should be true")
	}
	cell.SetString("This is not comment")
	if result := IsComment(cell); result != false {
		t.Error("'This is not comment' should be false")
	}
}

func TestCamelToSnake(t *testing.T) {
	if s := CamelToSnake("CamelCase"); s != "camel_case" {
		t.Error("'CamelCase' should be 'camel_case' but", s)
	}
	if s := CamelToSnake("getHTTPResponseCode"); s != "get_http_response_code" {
		t.Error("'getHTTPResponseCode' should be 'get_http_response_code' but", s)
	}
}

func TestExportFile(t *testing.T) {
	ExportFile(filename)
}

func TestExportCsvFile(t *testing.T) {
	xlsx_file, _ := xlsx.OpenFile(filename)
	ExportCsvFile(xlsx_file)
}

func TestExportJsonFile(t *testing.T) {
	xlsx_file, _ := xlsx.OpenFile(filename)
	ExportJsonFile(xlsx_file)
}

func TestExportKeyValueFile(t *testing.T) {
	xlsx_file, _ := xlsx.OpenFile(filename)
	ExportKeyValueFile(xlsx_file)
}

func TestExportSqlFile(t *testing.T) {
	xlsx_file, _ := xlsx.OpenFile(filename)
	ExportSqlFile(xlsx_file, filename)
}

func TestExportSqlAsJsonFile(t *testing.T) {
	xlsx_file, _ := xlsx.OpenFile(filename)
	ExportSqlAsJsonFile(xlsx_file)
}

func BenchmarkExportFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		ExportFile(filename)
	}
}

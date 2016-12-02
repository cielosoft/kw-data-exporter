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

func TestExportCsvFile(t *testing.T) {
	ExportCsvFile(filename)
	ExportCsvFile("Wrong filename")
}

func TestExportJsonFile(t *testing.T) {
	ExportJsonFile(filename)
	ExportJsonFile("Wrong filename")
}

func TestExportKeyValueFile(t *testing.T) {
	ExportKeyValueFile(filename)
	ExportKeyValueFile("Wrong filename")
}

func TestExportSqlFile(t *testing.T) {
	ExportSqlFile(filename)
	ExportSqlFile("Wrong filename")
}

func TestExportSqlAsJsonFile(t *testing.T) {
	ExportSqlAsJsonFile(filename)
	ExportSqlAsJsonFile("Wrong filename")
}

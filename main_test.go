package main

import "testing"

const filename string = "sample.xlsx"

func TestCamelToSnake(t *testing.T) {
	if s := CamelToSnake("CamelCase"); s != "camel_case" {
		t.Error("CamelCase should be camel_case but", s)
	}
	if s := CamelToSnake("getHTTPResponseCode"); s != "get_http_response_code" {
		t.Error("getHTTPResponseCode should be get_http_response_code but", s)
	}
}

func TestExportFile(t *testing.T) {
	ExportFile(filename)
}

func BenchmarkExportFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		ExportFile(filename)
	}
}

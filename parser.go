package main

import "github.com/tealeg/xlsx"
import "strings"
import "unicode"

type FieldInfo struct {
	col   int
	ftype string
	fname string
}

type Header struct {
	name         string
	exec         string
	fieldList    []FieldInfo
	csvFieldList []FieldInfo
}

func TrimString(cell *xlsx.Cell) string {
	s, _ := cell.String()
	// return strings.TrimSpace(s)
	return strings.Trim(s, "# ")
}

func IsComment(cell *xlsx.Cell) bool {
	s, _ := cell.String()
	return strings.HasPrefix(s, "#")
}

func CamelToSnake(str string) string {
	runes := []rune(str)
	length := len(runes)

	var out []rune
	for i := 0; i < length; i++ {
		if i > 0 && unicode.IsUpper(runes[i]) && ((i+1 < length && unicode.IsLower(runes[i+1])) || unicode.IsLower(runes[i-1])) {
			out = append(out, '_')
		}
		out = append(out, unicode.ToLower(runes[i]))
	}

	return string(out)
}

func ReadHeader(sheet *xlsx.Sheet) (header Header) {
	// 이름, 서버용 필드, 클라이언트용 필드, 형식 이렇게 최소 4줄이 필요
	if sheet.MaxRow < 4 || sheet.MaxCol < 2 {
		return
	}

	header.exec = strings.ToUpper(TrimString(sheet.Cell(0, 0)))
	header.name = TrimString(sheet.Cell(0, 1))

	// fieldList
	for col := 0; col < sheet.MaxCol; col++ {
		fname := TrimString(sheet.Cell(1, col))
		if len(fname) == 0 {
			continue
		}
		ftype := TrimString(sheet.Cell(3, col))

		header.fieldList = append(header.fieldList, FieldInfo{col: col, fname: fname, ftype: ftype})
	}
	// csvFieldList
	for col := 0; col < sheet.MaxCol; col++ {
		fname := TrimString(sheet.Cell(2, col))
		if len(fname) == 0 {
			continue
		}
		ftype := TrimString(sheet.Cell(3, col))

		header.csvFieldList = append(header.csvFieldList, FieldInfo{col: col, fname: fname, ftype: ftype})
	}
	return
}

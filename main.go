package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"path"
	"strconv"
	"strings"
	"unicode"
)

// 전역 변수
var USE_CSV bool
var USE_JSON bool
var USE_SQL bool

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

func CamelToSnake(in string) string {
	runes := []rune(in)
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
	if sheet.MaxRow < 4 || sheet.MaxCol < 2 {
		return
	}
	for i := 0; i < 4; i++ {
		if IsComment(sheet.Cell(i, 0)) == false {
			return
		}
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

func ExportCSVFile(xlsx_file *xlsx.File) {
	for _, sheet := range xlsx_file.Sheets {
		header := ReadHeader(sheet)
		if len(header.csvFieldList) == 0 {
			continue
		}

		var count int = 0
		var buffer string

		var columns []string
		for _, field := range header.csvFieldList {
			columns = append(columns, field.fname)
		}

		buffer += fmt.Sprintln(strings.Join(columns, "\t") + "\r")

		for r := 1; r < sheet.MaxRow; r++ {
			if IsComment(sheet.Cell(r, 0)) {
				continue
			}

			var values []string
			for _, field := range header.csvFieldList {
				cell := sheet.Cell(r, field.col)
				// 비어 있다
				if len(TrimString(cell)) == 0 {
					// fmt.Printf("row: %d, col: %d, %s\n\n", r, field.col, field.fname)
					continue
				}

				switch field.ftype {
				case "string":
					values = append(values, TrimString(cell))
				case "float":
					v, e := cell.Float()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					values = append(values, strconv.FormatFloat(v, 'f', -1, 32))
				case "auto":
					switch cell.Type() {
					case xlsx.CellTypeFormula, xlsx.CellTypeNumeric:
						v, _ := cell.Float()
						values = append(values, strconv.FormatFloat(v, 'f', -1, 32))
					default:
						values = append(values, TrimString(cell))
					}
				default:
					values = append(values, cell.Value)
				}
			}

			if len(values) == len(header.csvFieldList) {
				buffer += fmt.Sprintln(strings.Join(values, "\t") + "\r")
				count++
			}
		}

		fn := sheet.Name + ".csv"
		if err := ioutil.WriteFile(fn, []byte(buffer), 0644); err == nil {
			fmt.Println("Exported", path.Base(fn), count)
		} else {
			fmt.Println(err)
		}
	}
}

func ExportJsonFile(xlsx_file *xlsx.File) {
	for _, sheet := range xlsx_file.Sheets {
		header := ReadHeader(sheet)
		if !strings.Contains(header.exec, "JSON") {
			continue
		}
		if len(header.name) == 0 || len(header.fieldList) == 0 {
			continue
		}

		var data []interface{}
		for r := 1; r < sheet.MaxRow; r++ {
			if IsComment(sheet.Cell(r, 0)) {
				continue
			}

			var doc = make(map[string]interface{})
			for _, field := range header.fieldList {
				cell := sheet.Cell(r, field.col)
				switch field.ftype {
				case "string":
					value := TrimString(cell)
					if len(value) == 0 {
						fmt.Errorf("row: %d, col: %d, Field is empty\n", r, field.col)
						continue
					}
					doc[field.fname] = value
				case "float":
					v, e := cell.Float()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					doc[field.fname] = v
				default:
					v, e := cell.Int()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					doc[field.fname] = v
				}
			}
			if len(doc) == len(header.fieldList) {
				data = append(data, doc)
			}
		}

		if buffer, err := json.Marshal(data); err != nil {
			fmt.Println(err)
		} else {
			fn := CamelToSnake(header.name) + ".json"
			if err := ioutil.WriteFile(fn, buffer, 0644); err == nil {
				fmt.Println("Exported", path.Base(fn), len(data))
			} else {
				fmt.Println(err)
			}
		}
	}
}

func ExportKeyValueFile(xlsx_file *xlsx.File) {
	for _, sheet := range xlsx_file.Sheets {
		header := ReadHeader(sheet)
		if !strings.Contains(header.exec, "KEYVALUE") {
			continue
		}
		if len(header.name) == 0 || len(header.fieldList) == 0 {
			continue
		}

		var data = make(map[string]interface{})
		for r := 1; r < sheet.MaxRow; r++ {
			if IsComment(sheet.Cell(r, 0)) {
				continue
			}

			var key string = ""
			var value interface{}
			for _, field := range header.fieldList {
				cell := sheet.Cell(r, field.col)
				switch field.fname {
				case "key":
					key = TrimString(cell)
				case "value":
					switch field.ftype {
					case "string":
						value = TrimString(cell)
					case "float":
						v, e := cell.Float()
						if e != nil {
							fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
							continue
						}
						value = v
					default:
						v, e := cell.Int()
						if e != nil {
							fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
							continue
						}
						value = v
					}

				}
			}
			data[key] = value
		}

		if buffer, err := json.Marshal(data); err != nil {
			fmt.Println(err)
		} else {
			fn := CamelToSnake(header.name) + ".json"
			if err := ioutil.WriteFile(fn, buffer, 0644); err == nil {
				fmt.Println("Exported", path.Base(fn), len(data))
			} else {
				fmt.Println(err)
			}
		}
	}
}

func ExportSQLFile(xlsx_file *xlsx.File, filename string) {
	var count int = 0
	var buffer string

	buffer += fmt.Sprintf("-- Auto generated by kw-data-expoter\n")
	buffer += fmt.Sprintf("-- Source: %s\n", path.Base(filename))

	for _, sheet := range xlsx_file.Sheets {
		header := ReadHeader(sheet)
		if !strings.Contains(header.exec, "SQL") {
			continue
		}
		if len(header.name) == 0 || len(header.fieldList) == 0 {
			continue
		}

		var columns []string
		for _, field := range header.fieldList {
			columns = append(columns, "`"+field.fname+"`")
		}

		var data []string
		for r := 1; r < sheet.MaxRow; r++ {
			if IsComment(sheet.Cell(r, 0)) {
				continue
			}

			var values []string
			for _, field := range header.fieldList {
				cell := sheet.Cell(r, field.col)
				switch field.ftype {
				case "string":
					v, e := cell.String()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					values = append(values, "'"+v+"'")
				case "float":
					v, e := cell.Float()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					values = append(values, strconv.FormatFloat(v, 'f', -1, 32))
				default:
					v, e := cell.Int()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					values = append(values, strconv.Itoa(v))
				}
			}
			if len(values) == len(header.fieldList) {
				data = append(data, fmt.Sprintf("(%s)", strings.Join(values, ",")))
				count++
			}
		}

		buffer += fmt.Sprintf("\n-- Sheet: %s %d row(s)\n", sheet.Name, len(data))
		buffer += fmt.Sprintf("DELETE FROM `%s`;\n", header.name)
		buffer += fmt.Sprintf("INSERT INTO `%s` (%s) VALUES\n", header.name, strings.Join(columns, ","))
		buffer += strings.Join(data, ",\n") + ";\n"
	}

	if count > 0 {
		basename := path.Base(filename)
		ext := path.Ext(filename)
		fn := basename[0:len(basename)-len(ext)] + ".sql"
		if err := ioutil.WriteFile(fn, []byte(buffer), 0644); err == nil {
			fmt.Println("Exported", path.Base(fn), count)
		} else {
			fmt.Println(err)
		}
	}
}

func ExportFile(filename string) {
	xlsx_file, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Errorf("OpenFile Error:", filename)
		return
	}

	if USE_CSV {
		ExportCSVFile(xlsx_file)
	}

	if USE_JSON {
		ExportJsonFile(xlsx_file)
		ExportKeyValueFile(xlsx_file)
	}
	if USE_SQL {
		ExportSQLFile(xlsx_file, filename)
	}
}

func main() {
	no_csv := flag.Bool("no-csv", false, "csv 형식 사용 안함")
	flag_all := flag.Bool("all", false, "모든 형식 사용")
	flag.BoolVar(&USE_JSON, "json", false, "json 형식 사용")
	flag.BoolVar(&USE_SQL, "sql", false, "sql 형식 사용")
	flag.Parse()

	USE_CSV = !(*no_csv)
	if *flag_all {
		USE_JSON = true
		USE_SQL = true
	}

	// 기본은 현재 디렉토리
	var target string = "."
	if flag.NArg() > 0 {
		target = flag.Arg(0)
	}

	fileinfo, err := os.Stat(target)
	if os.IsNotExist(err) {
		return
	}

	if fileinfo.IsDir() {
		if files, err := ioutil.ReadDir(target); err == nil {
			for _, file := range files {
				if path.Ext(file.Name()) == ".xlsx" && !strings.HasPrefix(file.Name(), "~") {
					ExportFile(path.Join(target, file.Name()))
				}
			}
		}
	} else {
		ExportFile(target)
	}
}

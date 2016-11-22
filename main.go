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
)

// 전역 변수
var USE_CSV bool
var jsonflag bool
var sqlflag bool
var sqlasjson bool

func ExportJson(sheet *xlsx.Sheet, fields []FieldInfo) (data []interface{}) {
	for r := 1; r < sheet.MaxRow; r++ {
		if IsComment(sheet.Cell(r, 0)) {
			continue
		}

		var node = make(map[string]interface{})
		for _, field := range fields {
			cell := sheet.Cell(r, field.col)
			if len(cell.Value) == 0 {
				continue
			}

			switch field.ftype {
			case "string":
				value := TrimString(cell)
				node[field.fname] = value
			case "float":
				v, e := cell.Float()
				if e != nil {
					fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
					continue
				}
				node[field.fname] = v
			default:
				v, e := cell.Int()
				if e != nil {
					fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
					continue
				}
				node[field.fname] = v
			}
		}
		if len(node) == len(fields) {
			data = append(data, node)
		}
	}
	return
}

func ExportCsvFile(xlsx_file *xlsx.File) {
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

		data := ExportJson(sheet, header.fieldList)
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

func ExportSqlAsJsonFile(xlsx_file *xlsx.File) {
	for _, sheet := range xlsx_file.Sheets {
		header := ReadHeader(sheet)
		if !strings.Contains(header.exec, "SQL") {
			continue
		}
		if len(header.name) == 0 || len(header.fieldList) == 0 {
			continue
		}

		data := ExportJson(sheet, header.fieldList)
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
			if len(key) > 0 {
				data[key] = value
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

func ExportSqlFile(xlsx_file *xlsx.File, filename string) {
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
		ExportCsvFile(xlsx_file)
	}

	if jsonflag {
		ExportJsonFile(xlsx_file)
		ExportKeyValueFile(xlsx_file)
	}
	if sqlflag {
		ExportSqlFile(xlsx_file, filename)
	}
	if sqlasjson {
		ExportSqlAsJsonFile(xlsx_file)
	}
}

func main() {
	no_csv := flag.Bool("no-csv", false, "Not using csv format. csv is default on")
	flag_all := flag.Bool("all", false, "Same as --json --sql")
	flag.BoolVar(&jsonflag, "json", false, "Using json format")
	flag.BoolVar(&sqlflag, "sql", false, "Using sql format")
	flag.BoolVar(&sqlasjson, "sqlasjson", false, "Using sql format but as json")
	flag.Parse()

	USE_CSV = !(*no_csv)
	if *flag_all {
		jsonflag = true
		sqlflag = true
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

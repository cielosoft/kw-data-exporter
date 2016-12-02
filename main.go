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

func ExportJson(sheet *xlsx.Sheet, fields []FieldInfo) (data []interface{}) {
	for r := 1; r < sheet.MaxRow; r++ {
		if IsComment(sheet.Cell(r, 0)) {
			continue
		}

		var node = make(map[string]interface{})
		for _, field := range fields {
			cell := sheet.Cell(r, field.col)

			switch field.ftype {
			case "string":
				value := TrimString(cell)
				node[field.fname] = value
			case "float":
				if len(cell.Value) > 0 {
					v, e := cell.Float()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					node[field.fname] = v
				}
			default:
				if len(cell.Value) > 0 {
					v, e := cell.Int()
					if e != nil {
						fmt.Errorf("row: %d, col: %d, %s\n", r, field.col, e)
						continue
					}
					node[field.fname] = v
				}
			}
		}
		if len(node) == len(fields) {
			data = append(data, node)
		}
	}
	return
}

func ExportCsvFile(filename string) {
	xlsx_file, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Errorf("OpenFile Error:", filename)
		return
	}

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

		savefile := sheet.Name + ".csv"
		if err := ioutil.WriteFile(savefile, []byte(buffer), 0644); err == nil {
			fmt.Println("Exported", path.Base(filename), savefile, count)
		} else {
			fmt.Println(err)
		}
	}
}

func ExportJsonFile(filename string) {
	xlsx_file, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Errorf("OpenFile Error:", filename)
		return
	}

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
			savefile := CamelToSnake(header.name) + ".json"
			if err := ioutil.WriteFile(savefile, buffer, 0644); err == nil {
				fmt.Println("Exported", path.Base(filename), savefile, len(data))
			} else {
				fmt.Println(err)
			}
		}
	}
}

func ExportSqlAsJsonFile(filename string) {
	xlsx_file, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Errorf("OpenFile Error:", filename)
		return
	}

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
			savefile := CamelToSnake(header.name) + ".json"
			if err := ioutil.WriteFile(savefile, buffer, 0644); err == nil {
				fmt.Println("Exported", path.Base(filename), savefile, len(data))
			} else {
				fmt.Println(err)
			}
		}
	}
}

func ExportKeyValueFile(filename string) {
	xlsx_file, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Errorf("OpenFile Error:", filename)
		return
	}

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
			savefile := CamelToSnake(header.name) + ".json"
			if err := ioutil.WriteFile(savefile, buffer, 0644); err == nil {
				fmt.Println("Exported", path.Base(filename), savefile, len(data))
			} else {
				fmt.Println(err)
			}
		}
	}
}

func ExportSqlFile(filename string) {
	xlsx_file, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Errorf("OpenFile Error:", filename)
		return
	}

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
		savefile := basename[0:len(basename)-len(ext)] + ".sql"
		if err := ioutil.WriteFile(savefile, []byte(buffer), 0644); err == nil {
			fmt.Println("Exported", basename, savefile, count)
		} else {
			fmt.Println(err)
		}
	}
}

func main() {
	var flagNoCsv bool
	var flagJson bool
	var flagSql bool
	var flagSqlAsJson bool

	flagAll := flag.Bool("all", false, "Same as --json --sql")
	flag.BoolVar(&flagNoCsv, "no-csv", false, "Not using csv format. csv is default on")
	flag.BoolVar(&flagJson, "json", false, "Using json format")
	flag.BoolVar(&flagSql, "sql", false, "Using sql format")
	flag.BoolVar(&flagSqlAsJson, "sqlasjson", false, "Using sql format but as json")
	flag.Parse()

	if *flagAll {
		flagJson = true
		flagSql = true
	}

	// 기본은 현재 디렉토리
	var target string = "."
	if flag.NArg() > 0 {
		target = flag.Arg(0)
	}

	fileInfo, err := os.Stat(target)
	if err != nil {
		return
	}

	var fileList []string
	if fileInfo.IsDir() {
		if files, err := ioutil.ReadDir(target); err == nil {
			for _, file := range files {
				if path.Ext(file.Name()) == ".xlsx" && !strings.HasPrefix(file.Name(), "~") {
					fileList = append(fileList, path.Join(target, file.Name()))
				}
			}
		}
	} else {
		fileList = append(fileList, target)
	}

	for _, f := range fileList {
		if !flagNoCsv {
			ExportCsvFile(f)
		}
		if flagJson {
			ExportJsonFile(f)
			ExportKeyValueFile(f)
		}
		if flagSql {
			ExportSqlFile(f)
		}
		if flagSqlAsJson {
			ExportSqlAsJsonFile(f)
		}
	}
}

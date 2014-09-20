package main

import (
	"flag"
	"fmt"
	"xlsx"
	"os"
	"encoding/csv"
)


var xlsxPath = flag.String("f", "D:\\golang\\docxlsx\\XY_Army.xlsx", "Path to an XLSX file")
var csvPath = flag.String("c", "D:\\golang\\doccsv\\XY_Army.csv", "Path to an CSV file")
var sheetIndex = flag.Int("i",0, "Index of sheet to convert, zero based")
var delimiter = flag.String("d", ",", "Delimiter to use between fields")

type Outputer func(s string)

type XLSX2CSVError struct {
	error string
}

func (e XLSX2CSVError) Error() string {
	return e.error
}

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int,csvPath string, outputf Outputer) error {
	var xlFile *xlsx.File
	var error error
	var sheetLen int
	var rowString string

	xlFile, error = xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen = len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		e := new(XLSX2CSVError)
		e.error = "This XLSX file contains no sheets.\n"
		return *e
	case sheetIndex >= sheetLen:
		e := new(XLSX2CSVError)
		e.error = fmt.Sprintf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
		return *e
	}
	sheet := xlFile.Sheets[sheetIndex]

	file, error := os.Create(csvPath)
	if error!=nil {
		return error
	}

	defer file.Close()
	//file.WriteString("\xEF\xBB\xBF") // 写入UTF-8 BOM
	writer := csv.NewWriter(file)
	var record []string
	for _, row := range sheet.Rows {
		rowString = ""
		if row != nil {
			for cellIndex, cell := range row.Cells {
				record =append(record,cell.String())
				if cellIndex > 0 {
					rowString = fmt.Sprintf("%s%s%s", rowString, *delimiter, cell.String())
				} else {
					rowString = fmt.Sprintf("%s", cell.String())
				}
			}
			writer.Write(record)
			record=nil
			rowString = fmt.Sprintf("%s\n", rowString)
			//outputf(rowString)

		}
	}
	writer.Flush()
	return nil
}


func main() {
	flag.Parse()
	var error error
	error = generateCSVFromXLSXFile(*xlsxPath, *sheetIndex,*csvPath, func(s string) { fmt.Printf("%s", s) })
	if error != nil {
		fmt.Printf(error.Error())
		return
	}
}

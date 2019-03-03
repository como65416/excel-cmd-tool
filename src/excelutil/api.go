package excelutil

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

/**
 * get excel all sheet names
 */
func GetExcelSheetNames(filename string) ([]string, error) {
	xlsx, err := excelize.OpenFile(filename)
	var sheetNames = []string{}
	if err != nil {
		return sheetNames, err
	}
	for _, name := range xlsx.GetSheetMap() {
		sheetNames = append(sheetNames, name)
	}
	return sheetNames, nil
}

/**
 * get excel sheet data
 */
func ReadExcelContents(filepath string, sheetName string) ([][]string, error) {
	xlsx, err := excelize.OpenFile(filepath)
	var contents = [][]string{}
	if err != nil {
		fmt.Println(err)
		return [][]string{}, err
	}

	rows := xlsx.GetRows(sheetName)
	for _, row := range rows {
		var row_datas = []string{}
		for _, colCell := range row {
			row_datas = append(row_datas, colCell)
		}
		contents = append(contents, row_datas)
	}
	return contents, nil
}

func WriteAsExcel(filepath string, content [][]string) error {
	sheetRowLimit := 1000000
	datas_chunk := [][][]string{}
	for i := 0; i < len(content); i += sheetRowLimit {
		end := i + sheetRowLimit

		if end > len(content) {
			end = len(content)
		}

		datas_chunk = append(datas_chunk, content[i:end])
	}

	xlsx := excelize.NewFile()
	for sheetIndex, chunk := range datas_chunk {
		sheetName := "Sheet" + strconv.Itoa(sheetIndex)
		sheet := xlsx.NewSheet(sheetName)
		xlsx.SetActiveSheet(sheet)
		for i, row_data := range chunk {
			for j, cell_data := range row_data {
				var axios = ToExcelAxis(j+1, i+1)
				xlsx.SetCellValue(sheetName, axios, cell_data)
			}
			xlsx.Save()
		}
	}
	err := xlsx.SaveAs(filepath)
	return err
}

/**
 * get cell axis (column and row value start from 1)
 */
func ToExcelAxis(col int, row int) string {
	var axis = ""
	col -= 1
	for {
		axis = string(col%26+65) + axis
		col = col / 26

		if col == 0 {
			break
		}
	}
	axis = axis + strconv.Itoa(row)
	return axis
}

package excelutil

import (
	"fmt"

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

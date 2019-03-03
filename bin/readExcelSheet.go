package main

import (
	"excel-cmd-tool/src/excelutil"
	"fmt"
	"os"
	"path"
	"strings"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Command example : " + path.Base(os.Args[0]) + " input.xlsx sheetName")
		return
	}

	var filename = os.Args[1]
	var sheetName = os.Args[2]
	var contents, _ = excelutil.ReadExcelContents(filename, sheetName)

	for _, row := range contents {
		fmt.Println(strings.Join(row, "\t"))
	}
}

package main

import (
	"excel-cmd-tool/src/excelutil"
	"fmt"
	"os"
	"path"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Command example : " + path.Base(os.Args[0]) + " input.xlsx")
		return
	}

	var filename = os.Args[1]
	var sheetNames, err = excelutil.GetExcelSheetNames(filename)
	if err != nil {
		fmt.Println("Read file error : ")
		fmt.Println(err)
	}

	for _, name := range sheetNames {
		fmt.Println(name)
	}
}

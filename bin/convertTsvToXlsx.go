package main

import (
	"excel-cmd-tool/src/excelutil"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"strings"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Command example : " + path.Base(os.Args[0]) + " input.tsv output.xlsx")
		return
	}

	var tsv_filename = os.Args[1]
	var xlsx_filename = os.Args[2]

	data, _ := ioutil.ReadFile(tsv_filename)
	var content = string(data)
	var lines = strings.Split(content, "\n")
	var datas = [][]string{}
	for _, line := range lines {
		datas = append(datas, strings.Split(line, "\t"))
	}

	var err = excelutil.WriteAsExcel(xlsx_filename, datas)
	if err != nil {
		fmt.Println("Error :")
		fmt.Println(err)
	}
}

package main

import (
    "fmt"
	//"io/ioutil"
    "path/filepath"
    "github.com/360EntSecGroup-Skylar/excelize"
)


func main() {
    filesName, _ := filepath.Glob("./*.xlsx");
	fmt.Println(filesName)
	for i:=0;i<len(filesName);i++ {
		xlsx, err := excelize.OpenFile(filesName[i])
		if err != nil {
			fmt.Println("打开文件错误",filesName[i],err)
			return
		}
		
		cell := xlsx.GetCellValue("Sheet1", "B2")
		fmt.Println(cell)
		
		rows := xlsx.GetRows("Sheet1")
		for _, row := range rows {
			for _, colCell := range row {
				fmt.Print(colCell, "\t")
			}
			fmt.Println()
		}
	}
}
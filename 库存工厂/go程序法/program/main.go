package main

import (
	"fmt"
	//"io/ioutil"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func basename(s string) string {
	slash := strings.LastIndex(s, "\\") // -1 if "/" not found
	s = s[slash+1:]
	if dot := strings.LastIndex(s, "."); dot >= 0 {
		s = s[:dot]
	}
	return s
}

func main() {
	matchXlsx, _ := excelize.OpenFile("../match.xlsx")
	matchRows := matchXlsx.GetRows("Sheet1")
	fmt.Println("已登记商业公司个数：", len(matchRows)-1)
	fmt.Println("   ")
	count := 1    //记录当行已汇总的行数
	icount := 0   //计数导入了几列
	tmpCount := 0 //计数以判断是否在match表中有对应商业
	collectFile := excelize.NewFile()

	filesName, _ := filepath.Glob("../xlsx/*.xlsx")

	for i := 0; i < len(filesName); i++ {
		fmt.Println("正在处理文件：         ", basename(filesName[i])+".xlsx")
		xlsx, err := excelize.OpenFile(filesName[i])
		if err != nil {
			fmt.Println("打开文件错误", filesName[i], err)
			return
		}

		dataSheet := xlsx.GetSheetName(1) //获取文件中的第一张工作表
		if dataSheet == "" {
			fmt.Println("bad xlsx", i)
			continue
		}
		rows := xlsx.GetRows(dataSheet)
		rxa := len(rows)
		rya := len(rows[0])

		for m := 1; m < len(matchRows); m++ {
			if basename(filesName[i]) == matchRows[m][0] {
				fmt.Println("正在导入数据表：       ", dataSheet)
				for ry := 0; ry < rya; ry++ {
					switch rows[0][ry] {
					case matchRows[m][1]:
						for tmp := 1; tmp < rxa; tmp++ {
							collectFile.SetCellValue("Sheet1", "A"+strconv.Itoa(count+tmp-1), matchRows[m][5])
							collectFile.SetCellValue("Sheet1", "B"+strconv.Itoa(count+tmp-1), rows[tmp][ry])
							icount = icount + 1
						}

					case matchRows[m][2]:
						for tmp := 1; tmp < rxa; tmp++ {
							collectFile.SetCellValue("Sheet1", "C"+strconv.Itoa(count+tmp-1), rows[tmp][ry])
							icount = icount + 1
						}
					case matchRows[m][3]:
						for tmp := 1; tmp < rxa; tmp++ {
							collectFile.SetCellValue("Sheet1", "D"+strconv.Itoa(count+tmp-1), rows[tmp][ry])
							icount = icount + 1
						}

					case matchRows[m][4]:
						for tmp := 1; tmp < rxa; tmp++ {
							intCell, _ := strconv.ParseFloat(rows[tmp][ry], 64)
							collectFile.SetCellValue("Sheet1", "E"+strconv.Itoa(count+tmp-1), intCell)
							// collectFile.SetCellInt("Sheet1", axis string, value int)
							icount = icount + 1
						}
					}

					if icount == 4 {

						break
					}
				}
				if icount < 4 {
					fmt.Println("注意,有未找到，或未对应的列，注意,有未找到，或未对应的列，注意,有未找到，或未对应的列")
				}
				icount = 1
				count = count + rxa - 1
				fmt.Println("从此表中导入数据:      ", rxa-1, "行")
				fmt.Println("处理完成")
				fmt.Println("  ")
				break
			} else {
				tmpCount = tmpCount + 1
			}
		}
		if tmpCount == len(matchRows) {
			fmt.Println("未在对应表中找到此文件的对应项，未在对应表中找到此文件的对应项，未在对应表中找到此文件的对应项")
			fmt.Println("  ")
		}
		tmpCount = 1

	}

	// fmt.Println(collectFile.GetRows("Sheet1"))
	fmt.Println("所有文件处理完成，总共处理文件", len(filesName), "个，汇总数据", count-1, "行")

	collectFile.SaveAs("../汇总.xlsx")

}

/*
  @要汇总的表放在xlsx文件夹里，要求格式必须为.xlsx。
  @注意最好把表里的格式清除一下， 同时保证你要汇总的那个sheet是第一个sheet
  @汇总完了会在当前目录生成  汇总.xlsx
  excelize.go文件中的61行在打开文件时会调用readerfiles来读取文件，是用了readerall，即一下全部读入，这倒
       是没什么问题，但是在lib.go中有一个namespaceStrictToTransitional()函数，它在替换是如果文件很大，会花很多的时间


*/

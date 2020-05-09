package tools

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"path"
	"time"
)

type Translate struct {
	Path      string
	ExcelPath map[int]string
}

func (this *Translate) Run() {
	t := time.Now()
	files, _ := ioutil.ReadDir(this.Path)
	i := 0
	this.ExcelPath = make(map[int]string)
	// 获取文件并输出它们的名字
	for _, file := range files {
		//go fmt.Println(file.Name())
		// 判断是否为xlsx文件
		fileName := path.Base(this.Path + file.Name())
		if ".xlsx" == path.Ext(fileName) {
			this.ExcelPath[i] = this.Path + file.Name()
			i++
		}
	}
	// wg 如果在函数里传递过去是值传递 相当于复制了一份 起不到控制的效果
	wg.Add(len(this.ExcelPath))
	for _, filePath := range this.ExcelPath {
		go this.ReadExcel(filePath)

	}
	wg.Wait()
	fmt.Println("替换excel文件完成，所耗时间：", time.Since(t))
}

func (this *Translate) ReadExcel(filePath string) {
	defer wg.Done()
	file, e := excelize.OpenFile(filePath)
	if e != nil {
		fmt.Println("读取excel文件失败:", e)
		return
	}
	var excelMaps map[int]interface{}
	excelMaps = make(map[int]interface{})
	var writeFilePath string
	i := 0
	for _, name := range file.GetSheetMap() {
		// 获取全部单元格的值
		rows, e := file.GetRows(name)
		fmt.Println("打印当前rows的长度:", len(rows))
		if e != nil {
			fmt.Println("获取单元格的值失败:", e)
			return
		}
		for _, row := range rows {
			for col, colCell := range row {
				if colCell == "" {
					continue
				}
				fmt.Printf("col:%v colCell:%v\n", col, colCell)
				switch col {
				case 0:
					writeFilePath = colCell
				}
			}
			if len(row) >= 5 {
				excelMaps[i] = row
				i++
			}
			fmt.Printf("将%s存入map中\n", filePath)
		}
	}
	fmt.Printf("%s存储完毕\n", filePath)
	this.TransLate(&excelMaps, writeFilePath)
}

func (this *Translate) TransLate(excelMaps *map[int]interface{}, writeFilePath string) error {
	fmt.Println("这里打印excel的目录:", writeFilePath)
	var file *excelize.File
	var length int
	length = len(*excelMaps)
	file, err = excelize.OpenFile(writeFilePath)
	if err != nil {
		fmt.Println("打印翻译错误的err:", err)
		return err
	}
	for _, excelInterface := range *excelMaps {
		excelMap := excelInterface.([]string)
		// 0为文件路径 1为工作表名称 2为坐标 3为原来内容 4为替换后的内容
		if excelMap[4] != "" {
			file.SetCellValue(excelMap[1], excelMap[2], excelMap[4])
			if err != nil {
				return err
			}
		}
		length--
		fmt.Printf("%s还剩%d\n", excelMap[0], length)
	}
	if err := file.Save(); err != nil {
		println(err.Error())
	}
	return err
}

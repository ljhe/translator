package tools

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"os"
	"path"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode"
)

var (
	wg                sync.WaitGroup
	err               error
	unUsefulRowLength = 3                          // 这里设置不需翻译的前几行数据
	skipSheetName     = []string{"Sheet", "sheet"} // 这里设置需要跳过的工作表名称
	fieldLine         = 2                          // 这里设置字段所在的行数 计数从0开始
)

type PickUp struct {
	Path      string
	ExcelPath map[int]string
}

func (this *PickUp) Run() {
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
		// 该程序是按支持并发设计的 但是问了Excelize插件的制作者 该插件是不支持并发的
		// 所以这里保留了并发程序 但是并达不到并发插入效果 可以并发的将数据存入内存
		go this.ReadExcel(filePath)

	}
	wg.Wait()
	fmt.Println("读取excel文件完成，所耗时间：", time.Since(t))
}

// 读取Excel文档
func (this *PickUp) ReadExcel(filePath string) {
	defer wg.Done()
	var rowForExcel = 1
	splitNewFileName := strings.Split(filePath, `\`)
	fileName := splitNewFileName[len(splitNewFileName)-1]
	// 记录坐标
	var excelMaps map[int]interface{}
	excelMaps = make(map[int]interface{})
	file, e := excelize.OpenFile(filePath)
	if e != nil {
		fmt.Println("读取excel文件失败:", e)
		return
	}

	for _, name := range file.GetSheetMap() {
		var colSince []int
		// sheet名全为英文的需要转出
		if IsChinese(name) {
			continue
		}
		fmt.Println("name:", name)
		flag := false
		for _, skip := range skipSheetName {
			if strings.Contains(name, skip) {
				flag = true
			}
		}
		if flag {
			continue
		}
		// 获取全部单元格的值
		rows, e := file.GetRows(name)
		if e != nil {
			fmt.Println("获取单元格的值失败:", e)
			return
		}
		// 第一行记录数据类型
		if rows == nil || len(rows) <= fieldLine {
			continue
		}
		i := 0
		for k, v := range rows[fieldLine] {
			if v != "" && IsLetter(v) {
				colSince = append(colSince, k)
				i++
			}
		}
		fmt.Println("colSince:", len(colSince))
		for key, row := range rows {
			// excel表格前几行是无用数据 不需要读取
			if key < unUsefulRowLength {
				continue
			}
			if len(row) == 0 {
				continue
			}
			rowLength := len(row)
			for _, v1 := range colSince {
				if v1 >= rowLength {
					continue
				}
				if IsChinese(row[v1]) {
					//索引转化为坐标
					cellName, err := excelize.CoordinatesToCellName(v1+1, key+1)
					if err != nil {
						fmt.Println("坐标转换失败:", e)
						return
					}
					excelMaps[rowForExcel] = map[string]interface{}{
						"filePath":    filePath,
						"name":        name,
						"cellName":    cellName,
						"colCell":     row[v1],
						"rowForExcel": rowForExcel,
					}
					//this.WriteExcel(filePath, fileName, name, cellName, colCell, rowForExcel, writeFilePath, writeFile)
					rowForExcel++
				}
			}
			fmt.Println(row)
		}
	}
	this.WriteExcel(&excelMaps, fileName)
}

// 创建一个Excel文档
func (this *PickUp) WriteExcel(excelMaps *map[int]interface{}, fileName string) error {
	var sheetName = "Sheet1"
	var filePath, sheet, cellName, value string
	var rowForExcel int
	file := excelize.NewFile()
	index := file.NewSheet(sheetName)
	file.SetActiveSheet(index)
	for _, excelMapInterface := range *excelMaps {
		excelMap := excelMapInterface.(map[string]interface{})
		rowForExcel = excelMap["rowForExcel"].(int)
		sheet = excelMap["name"].(string)
		cellName = excelMap["cellName"].(string)
		value = excelMap["colCell"].(string)
		filePath = excelMap["filePath"].(string)
		err = file.SetSheetRow(sheetName, "A"+strconv.Itoa(rowForExcel), &[]interface{}{filePath, sheet, cellName, value})
		fmt.Println("开始转化:", value)
	}
	// 判断文件所需转化是否为空
	if len(*excelMaps) != 0 {
		// 根据指定路径保存文件
		if err = file.SaveAs(fileName); err != nil {
			fmt.Println("写入Excel文件失败!", err)
		}
	}
	fmt.Println("rowForExcel:", rowForExcel)
	return err
}

// 判断是否为汉字
func IsChinese(r string) bool {
	for _, v := range []rune(r) {
		if unicode.Is(unicode.Han, v) {
			return true
		}
	}
	return false
}

// 判断是否为字母
func IsLetter(r string) bool {
	matched, _ := regexp.MatchString("([a-zA-Z])+([0-9])*", r)
	if matched {
		return true
	}
	return false
}

// 判断所给路径文件/文件夹是否存在
func PathExists(path string) bool {
	// os.Stat获取文件信息
	_, err := os.Stat(path)
	if err != nil {
		if os.IsExist(err) {
			return true
		}
		return false
	}
	return true
}

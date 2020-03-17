package tools

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"testing"
)

func TestIsChinese(t *testing.T) {
	a := "你"
	fmt.Println(IsChinese(a))
	a = "111你"
	fmt.Println(IsChinese(a))
	a = "23asdasd213"
	fmt.Println(IsChinese(a))
	a = "你asdassda"
	fmt.Println(IsChinese(a))
}

func TestPathExists(t *testing.T) {
	fmt.Println(PathExists(`F:\code\src\translator`))
	fmt.Println(PathExists(`F:\code\src\translator\game.xlsx`))
	fmt.Println(PathExists(`F:\code\src\translator\notExists.xlsx`))
	dir, _ := os.Getwd()
	println(dir)
}

func TestPickUp_WriteExcel(t *testing.T) {
	file, err := excelize.OpenFile(`F:\code\src\translator\test\guild.xlsx`)
	if err != nil {
		fmt.Println("open file error:", err)
	}
	file.SetCellValue("guildDaily", "D4", "lala1")
	file.SetCellValue("guildDaily", "D5", "lala2")
	file.SetCellValue("buildings", "F4", "lala3")
	if err := file.Save(); err != nil {
		println(err.Error())
	}
}

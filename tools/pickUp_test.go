package tools

import (
	"fmt"
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

func TestIsLetter(t *testing.T) {
	fmt.Println(IsLetter("asdasd"))
	fmt.Println(IsLetter("asdasd1"))
	fmt.Println(IsLetter("你说"))
	fmt.Println(IsLetter("1"))
	fmt.Println(IsLetter(""))
}

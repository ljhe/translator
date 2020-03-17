package main

import (
	"bufio"
	"flag"
	"fmt"
	"github.com/gpmgo/gopm/modules/base"
	"os"
	"translator/tools"
)

var (
	// 需要提取汉字的Excel文件目录
	dirPath = flag.String("dirPath", `F:\code\src\translator\test\`, "target dir path")
	// 需要翻译的Excel文件目录
	translatePath = flag.String("translatePath", `F:\code\src\translator\`, "translate dir path")
)

func main() {
	input := bufio.NewScanner(os.Stdin)
	fmt.Println("Please input the number for program")
	fmt.Println("1.Extract Chinese characters")
	fmt.Println("2.Translation of foreign languages")
	var status string
	for input.Scan() {
		line := input.Text()
		if line == "1" || line == "2" {
			status = line
			break
		}
		fmt.Println("error param, please try again")
	}

	if status == "1" {
		if !base.IsDir(*dirPath) {
			panic("path is error!")
		}
		pickUp := tools.PickUp{Path: *dirPath}
		pickUp.Run()
	} else if status == "2" {
		if !base.IsDir(*translatePath) {
			panic("path is error!")
		}
		translate := tools.Translate{Path: *translatePath}
		translate.Run()
	}
}

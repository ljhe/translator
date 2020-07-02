package main

import (
	"bufio"
	"fmt"
	"github.com/gpmgo/gopm/modules/base"
	"newTranslator/tools"
	"os"
)

func main() {
	input := bufio.NewScanner(os.Stdin)
	fmt.Println("请选择功能（路径不能包含中文）：\n1.提取目标路径中Excel中的汉字\n2.将翻译过后的内容替换到原Excel文件，" +
		"请确保原Excel路径及文件存在\n0.退出")
	var status string
	var path string
	for input.Scan() {
		line := input.Text()
		if line == "1" || line == "2" {
			status = line
			fmt.Println("请输入目标Excel的路径：")
			break
		}
		fmt.Println("error param, please try again")
		if line == "0" {
			return
		}
	}

	for input.Scan() {
		path = input.Text()
		if path == "0" {
			return
		}
		choiceProgram(path, status)
	}
}

func choiceProgram(path, status string) {
	switch status {
	case "1":
		if base.IsDir(path) {
			pickUp := tools.PickUp{Path: checkPath(path)}
			pickUp.Run()
			fmt.Println("``````````````````任务完成 输入0退出``````````````````")
		} else {
			fmt.Println("error path, please try again")
		}
	case "2":
		if base.IsDir(path) {
			translate := tools.Translate{Path: checkPath(path)}
			translate.Run()
			fmt.Println("``````````````````任务完成 输入0退出``````````````````")
		} else {
			fmt.Println("error path, please try again")
		}
	case "0":
		return
	}
}

func checkPath(path string) string {
	if string(path[len(path)-1]) != `\` {
		path = path + `\`
	}
	return path
}

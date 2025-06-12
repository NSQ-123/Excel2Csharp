package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/Excel2Csharp/exceltool"
)

func main() {

	fmt.Println("Excel to CSV Converter")

	input_directory := "./asset/excel" // 当前目录下的asset/excel
	output_directory := "./asset/csv"  // 当前目录下的asset/csv

	// 打印输入输出目录的绝对路径
	inputAbs, _ := filepath.Abs(input_directory)
	outputAbs, _ := filepath.Abs(output_directory)
	fmt.Printf("输入目录绝对路径: %s\n", inputAbs)
	fmt.Printf("输出目录绝对路径: %s\n", outputAbs)

	// 检查输入目录是否存在
	if _, err := os.Stat(input_directory); os.IsNotExist(err) {
		fmt.Printf("输入目录不存在: %s\n", input_directory)
		return
	}

	// 检查输出目录是否存在，如果不存在则创建
	if _, err := os.Stat(output_directory); os.IsNotExist(err) {
		err := os.MkdirAll(output_directory, 0755)
		if err != nil {
			fmt.Printf("创建输出目录失败: %v\n", err)
			return
		}
	}

	//调用Excel2Csv.go中的ConvertAll函数
	err := exceltool.ConvertAll(input_directory, output_directory)
	if err != nil {
		fmt.Println("Error:", err)
	}
}

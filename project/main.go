package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/Excel2Csharp/exceltool"
)

func main() {

	fmt.Println("Excel to CSV Converter")

	input_excel := "./asset/excel"    // 当前目录下的asset/excel
	output_csv := "./asset/csv"       // 当前目录下的asset/csv
	output_csharp := "./asset/csharp" //当前目录下的asset/csharp

	// 打印输入输出目录的绝对路径
	inputExcelAbs, _ := filepath.Abs(input_excel)
	outputCsvAbs, _ := filepath.Abs(output_csv)
	outputCsharpAbs, _ := filepath.Abs(output_csharp)

	fmt.Printf("excel输入目录绝对路径: %s\n", inputExcelAbs)
	fmt.Printf("csv输出目录绝对路径: %s\n", outputCsvAbs)
	fmt.Printf("csharp输出目录绝对路径: %s\n", outputCsharpAbs)

	// 检查输入目录是否存在
	if _, err := os.Stat(input_excel); os.IsNotExist(err) {
		fmt.Printf("输入目录不存在: %s\n", input_excel)
		return
	}

	// 检查输出目录是否存在，如果不存在则创建
	if _, err := os.Stat(output_csv); os.IsNotExist(err) {
		err := os.MkdirAll(output_csv, 0755)
		if err != nil {
			fmt.Printf("创建输出目录失败: %v\n", err)
			return
		}
	}

	// 检查输出目录是否存在，如果不存在则创建
	if _, err := os.Stat(output_csharp); os.IsNotExist(err) {
		err := os.MkdirAll(output_csharp, 0755)
		if err != nil {
			fmt.Printf("创建输出目录失败: %v\n", err)
			return
		}
	}

	// //调用Excel2Csv.go中的ConvertAll2Csv函数
	// err := exceltool.ConvertAll2Csv(input_excel, output_csv)
	// if err != nil {
	// 	fmt.Println("Error:", err)
	// }

	errC := exceltool.ConvertAll2Csharp(input_excel, output_csharp)
	if errC != nil {
		fmt.Println("Error:", errC)
	}

}

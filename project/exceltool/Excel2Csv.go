// Package exceltool 提供将Excel文件转换为CSV格式的功能
// 主要功能包括批量转换目录下所有xlsx文件为csv，支持读取指定的sheet和处理元数据
// 以及按规则导出数据到CSV文件
// 依赖库: github.com/xuri/excelize/v2
// 版本: v2.9.1
package exceltool

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

var _addBracket = false // 是否在 arr 类型的分组括号中添加括号，true 为添加括号，false 为不添加

// ConvertAll2Csv 批量转换目录下所有 xlsx 文件为 csv
func ConvertAll2Csv(inputDir, outputDir string) error {
	files, err := filepath.Glob(filepath.Join(inputDir, "*.xlsx"))
	if err != nil {
		return fmt.Errorf("查找xlsx文件失败: %w", err)
	}
	if len(files) == 0 {
		return fmt.Errorf("目录下没有xlsx文件: %s", inputDir)
	}
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		if err := os.MkdirAll(outputDir, 0755); err != nil {
			return fmt.Errorf("创建输出目录失败: %w", err)
		}
	}
	for _, file := range files {
		base := filepath.Base(file)
		name := strings.TrimSuffix(base, filepath.Ext(base))
		csvPath := filepath.Join(outputDir, name+".csv")
		if err := Convert(file, csvPath); err != nil {
			return fmt.Errorf("转换文件失败 %s: %w", file, err)
		}
	}
	return nil
}

// Convert 单个xlsx转csv，仅读取第1和第2个sheet，按规则导出
func Convert(xlsxFile, csvFile string) error {
	f, err := excelize.OpenFile(xlsxFile)
	if err != nil {
		return fmt.Errorf("打开Excel失败: %w", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) < 2 {
		return fmt.Errorf("Excel文件至少需要2个sheet")
	}
	dataSheet := sheets[0]
	metaSheet := sheets[1]

	dataRows, err := f.GetRows(dataSheet)
	if err != nil {
		return fmt.Errorf("读取数据sheet失败: %w", err)
	}
	metaRows, err := f.GetRows(metaSheet)
	if err != nil {
		return fmt.Errorf("读取元数据sheet失败: %w", err)
	}
	if len(metaRows) < 3 {
		return fmt.Errorf("元数据sheet至少3行")
	}

	csvF, err := os.Create(csvFile)
	if err != nil {
		return fmt.Errorf("创建CSV文件失败: %w", err)
	}
	defer csvF.Close()
	writer := csv.NewWriter(csvF)
	defer writer.Flush()

	for i, row := range dataRows {
		if i == 0 {
			continue // 跳过第一行字段名
		}
		if len(row) == 0 || strings.TrimSpace(row[0]) == "" {
			continue // 第一列为空跳过
		}
		var csvLine []string
		for j := 0; j < len(metaRows[0]); j++ {
			fieldType := getMetaCell(metaRows, 1, j)
			fieldUsage := getMetaCell(metaRows, 2, j)
			if fieldType == "" || fieldUsage == "" || strings.Contains(strings.ToLower(fieldUsage), "n") {
				continue
			}
			if !strings.Contains(strings.ToLower(fieldUsage), "c") {
				continue // 只导出客户端字段
			}
			cellStr := ""
			if j < len(row) {
				cellStr = row[j]
			}
			//1. arr类型 添加括号
			//csvLine = append(csvLine, processCellValue(fieldType, cellStr, true))
			//2. arr类型 不添加括号
			csvLine = append(csvLine, processCellValue(fieldType, cellStr, _addBracket))
		}
		if len(csvLine) > 0 {
			writer.Write(csvLine)
		}
	}
	return nil
}

// getMetaCell 安全获取元数据单元格内容
func getMetaCell(metaRows [][]string, row, col int) string {
	if row < len(metaRows) && col < len(metaRows[row]) {
		return metaRows[row][col]
	}
	return ""
}

// 处理 arr 类型分组括号的通用方法
func processArrGroups(cellStr string, addBracket bool) string {
	groups := strings.Split(cellStr, ";")
	for i, group := range groups {
		group = strings.TrimSpace(group)
		if group == "" {
			continue
		}
		if addBracket {
			// 如果没有括号，加上
			if !strings.HasPrefix(group, "(") && !strings.HasSuffix(group, ")") {
				group = "(" + group + ")"
			}
		} else {
			// 去除已有括号
			group = strings.TrimPrefix(group, "(")
			group = strings.TrimSuffix(group, ")")
		}
		groups[i] = group
	}
	return strings.Join(groups, ";")
}

// processCellValue 按类型处理单元格内容（可扩展）
func processCellValue(fieldType, cellStr string, addBracket bool) string {
	//fmt.Printf("Processing field type: %s, cell value: %s\n", fieldType, cellStr)
	lowerFieldType := strings.ToLower(fieldType)

	if strings.HasPrefix(lowerFieldType, "arr<") && strings.HasSuffix(lowerFieldType, ">") {
		// 这是 arr<...> 格式
		return processArrGroups(cellStr, addBracket)
	}

	switch strings.ToLower(lowerFieldType) {
	case "int":
		return cellStr
	case "float", "double":
		return cellStr
	case "long":
		return cellStr
	case "string":
		return cellStr
	case "bool":
		return cellStr
	case "datetime":
		return cellStr
	default:
		return cellStr
	}
}

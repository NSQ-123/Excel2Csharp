package exceltool

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ConvertAll 读取 inputDir 下所有 xlsx 文件，生成 C# 类文件到 outputDir
func ConvertAll2Csharp(inputDir, outputDir string) error {
	files, err := filepath.Glob(filepath.Join(inputDir, "*.xlsx"))
	if err != nil {
		return err
	}
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		if err := os.MkdirAll(outputDir, 0755); err != nil {
			return err
		}
	}
	var classNames []string
	for _, file := range files {
		fileName := strings.TrimSuffix(filepath.Base(file), ".xlsx")
		outputFile := filepath.Join(outputDir, fileName+".cs")
		className, err := ConvertToCsharp(file, outputFile)
		if err != nil {
			fmt.Printf("转换失败: %s, err: %v\n", file, err)
			continue
		}
		classNames = append(classNames, className)
	}
	// 生成 TableDataLoader
	if len(classNames) > 0 {
		err := createTableDataLoader(outputDir, classNames)
		if err != nil {
			fmt.Printf("生成 TableDataLoader 失败: %v\n", err)
		}
	}
	return nil
}

// createTableDataLoader 生成 TableDataLoader.cs 文件
func createTableDataLoader(outputDir string, classNames []string) error {
	const (
		indent0        = ""
		indent1        = "\t"
		indent2        = "\t\t"
		indent3        = "\t\t\t"
		NAME_SPACE     = "GameFramework.Table"
		AsyncOperation = "Task" // 可根据需要调整
	)
	var loaderBuilder strings.Builder
	loaderBuilder.WriteString(fmt.Sprintf("%susing System;\n", indent0))
	loaderBuilder.WriteString(fmt.Sprintf("%susing System.Collections.Generic;\n", indent0))
	loaderBuilder.WriteString("\n")
	loaderBuilder.WriteString(fmt.Sprintf("%snamespace %s\n%s{\n", indent0, NAME_SPACE, indent0))
	loaderBuilder.WriteString(fmt.Sprintf("%spublic class TableDataLoader\n%s{\n", indent1, indent1))
	loaderBuilder.WriteString(fmt.Sprintf("%spublic static async %s LoadAll()\n%s{\n", indent2, AsyncOperation, indent2))
	loaderBuilder.WriteString(fmt.Sprintf("%sList<%s> tasks = new();\n", indent3, AsyncOperation))

	for _, className := range classNames {
		// 传入 className 去掉前2个字符（T_）
		param := className
		if len(className) > 2 {
			param = className[2:]
		}
		loaderBuilder.WriteString(fmt.Sprintf("%stasks.Add(%s.LoadAll(\"%s\"));\n", indent3, className, param))
	}

	loaderBuilder.WriteString(fmt.Sprintf("%sawait %s.WhenAll(tasks);\n", indent3, AsyncOperation))
	loaderBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))
	loaderBuilder.WriteString(fmt.Sprintf("%s}\n", indent1))
	loaderBuilder.WriteString(fmt.Sprintf("%s}\n", indent0))
	outputFilePath := filepath.Join(outputDir, "TableDataLoader.cs")
	return os.WriteFile(outputFilePath, []byte(loaderBuilder.String()), 0644)
}

// ConvertToCsharp 读取 Excel 的第二个工作表，生成 C# 类文件
func ConvertToCsharp(excelPath, outputPath string) (string, error) {
	const (
		indent0    = ""
		indent1    = "\t"
		indent2    = "\t\t"
		indent3    = "\t\t\t"
		NAME_SPACE = "GameFramework.Table"
		Dictionary = "_dataMap"
	)
	fileName := strings.TrimSuffix(filepath.Base(excelPath), ".xlsx")
	className := "T_" + fileName

	f, err := excelize.OpenFile(excelPath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) < 2 {
		return "", fmt.Errorf("Excel 文件 %s 少于2个工作表", excelPath)
	}
	sheetName := sheets[1] // 第二个工作表

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}
	if len(rows) < 4 {
		return "", fmt.Errorf("工作表格式不正确，至少需要4行定义")
	}
	fieldNames := rows[0]
	fieldTypes := rows[1]
	usages := rows[2]
	descriptions := rows[3]

	var classBuilder strings.Builder
	classBuilder.WriteString(fmt.Sprintf("%susing System;\n", indent0))
	classBuilder.WriteString(fmt.Sprintf("%susing System.Collections.Generic;\n\n", indent0))
	classBuilder.WriteString(fmt.Sprintf("%snamespace %s\n%s{\n", indent0, NAME_SPACE, indent0))
	classBuilder.WriteString(fmt.Sprintf("%spublic partial class %s : ITable\n%s{\n", indent1, className, indent1))
	classBuilder.WriteString(fmt.Sprintf("%sprivate static readonly Dictionary<int, %s> %s = new Dictionary<int, %s>();\n", indent2, className, Dictionary, className))
	classBuilder.WriteString(fmt.Sprintf("%sprivate static List<%s> _dataList;\n\n", indent2, className))

	// 字段builder
	var fieldLoadBuilder strings.Builder
	// 嵌套类型builder
	var subClassBuilder strings.Builder
	//var idField string

	for i := 0; i < len(fieldNames); i++ {
		if i >= len(fieldTypes) || i >= len(usages) {
			continue
		}
		fieldName := fieldNames[i]
		fieldType := fieldTypes[i]
		usage := usages[i]
		description := ""
		if i < len(descriptions) {
			description = descriptions[i]
		}
		if !strings.Contains(strings.ToLower(usage), "c") {
			continue
		}
		if description != "" {
			classBuilder.WriteString(fmt.Sprintf("%s/// <summary>\n%s/// %s\n%s/// </summary>\n", indent2, indent2, description, indent2))
		}
		fieldNameGo := strings.Title(fieldName)
		csharpType := CorrectType(fieldType)
		// 检查是否为嵌套类型 arr<...>
		if strings.HasPrefix(strings.ToLower(fieldType), "arr<") && strings.HasSuffix(fieldType, ">") {
			arrType := "T_" + fieldNameGo
			classBuilder.WriteString(fmt.Sprintf("%spublic List<%s> %s { get; set; }\n", indent2, arrType, fieldNameGo))
			fieldLoadBuilder.WriteString(fmt.Sprintf("%sthis.%s = ConvertUtils.LoadArr<%s>(data[%d]);\n", indent3, fieldNameGo, arrType, i))
			// 生成子类
			ProcessArr(fieldType, arrType, &subClassBuilder, indent1, indent2, indent3)
		} else {
			classBuilder.WriteString(fmt.Sprintf("%spublic %s %s { get; set; }\n", indent2, csharpType, fieldNameGo))

			fieldLoadBuilder.WriteString(fmt.Sprintf("%sthis.%s = ConvertUtils.Get<%s>(data[%d]);\n", indent3, fieldNameGo, csharpType, i))
		}
	}

	// 生成静态方法 GetById
	classBuilder.WriteString("\n")
	classBuilder.WriteString(fmt.Sprintf("%spublic static %s GetById(int id)\n%s{\n", indent2, className, indent2))
	classBuilder.WriteString(fmt.Sprintf("%sif (%s.TryGetValue(id, out var value))\n%s{\n", indent3, Dictionary, indent3))
	classBuilder.WriteString(fmt.Sprintf("%sreturn value;\n", indent3+indent1))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%sreturn null;\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))

	// 生成静态方法 GetAll
	classBuilder.WriteString("\n")
	classBuilder.WriteString(fmt.Sprintf("%spublic static List<%s> GetAll()\n%s{\n", indent2, className, indent2))
	classBuilder.WriteString(fmt.Sprintf("%sif (_dataList == null)\n%s{\n", indent3, indent3))
	classBuilder.WriteString(fmt.Sprintf("%s_dataList = new List<%s>(%s.Values);\n", indent3+indent1, className, Dictionary))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%sreturn _dataList;\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))

	// 生成实例方法 Load
	classBuilder.WriteString("\n")
	classBuilder.WriteString(fmt.Sprintf("%spublic void Load(string[] data)\n%s{\n", indent2, indent2))
	classBuilder.WriteString(fieldLoadBuilder.String())
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))

	// 生成实例方法 GetId
	classBuilder.WriteString("\n")
	classBuilder.WriteString(fmt.Sprintf("%spublic int GetId()\n", indent2))
	classBuilder.WriteString(fmt.Sprintf("%s{\n", indent2))
	classBuilder.WriteString(fmt.Sprintf("%svar idProperty = this.GetType().GetProperty(\"ID\");\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%sif (idProperty != null)\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%s{\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%s\treturn (int)idProperty.GetValue(this);\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%sthrow new Exception($\"当前类 {{this.GetType().Name}}  未定义 ID 属性\");\n", indent3))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent1))
	classBuilder.WriteString("\n")

	// 如果存在 arr<...> 类型的字段，则生成对应的子类
	if subClassBuilder.Len() > 0 {
		classBuilder.WriteString("\n")
		classBuilder.WriteString(subClassBuilder.String())
	}
	// 结束命名空间
	classBuilder.WriteString(fmt.Sprintf("%s}\n", indent0))

	// 写入文件
	if err := os.WriteFile(outputPath, []byte(classBuilder.String()), 0644); err != nil {
		return "", err
	}
	fmt.Printf("生成 C# 类：%s\n", className)
	return className, nil
}

// ProcessArr 生成嵌套类型子类
func ProcessArr(fieldType, className string, subBuilder *strings.Builder, indent1, indent2, indent3 string) {
	// 这里只做简单处理，假设 arr<int> 或 arr<string> 或 arr<type1,type2,...>
	subBuilder.WriteString(fmt.Sprintf("%spublic partial class %s : ITable\n%s{\n", indent1, className, indent1))
	inner := fieldType[4 : len(fieldType)-1] // 去掉 arr< 和 >
	types := strings.Split(inner, ",")
	for i, t := range types {
		t = strings.TrimSpace(t)
		if strings.HasSuffix(t, "slice") {
			baseType := strings.TrimSuffix(t, "slice")
			baseType = strings.TrimSpace(baseType)
			if baseType == "" {
				baseType = "int"
			}
			goType := CorrectType(baseType)
			subBuilder.WriteString(fmt.Sprintf("%spublic List<%s> Args%d;\n", indent2, goType, i))
		} else {
			baseType := CorrectType(t)
			subBuilder.WriteString(fmt.Sprintf("%spublic %s Args%d;\n", indent2, baseType, i))
		}
	}
	// Load 方法
	subBuilder.WriteString("\n")
	subBuilder.WriteString(fmt.Sprintf("%spublic void Load(string[] data)\n%s{\n", indent2, indent2))
	for i, t := range types {
		t = strings.TrimSpace(t)
		if strings.HasSuffix(t, "slice") {
			baseType := strings.TrimSuffix(t, "slice")
			baseType = strings.TrimSpace(baseType)
			if baseType == "" {
				baseType = "int"
			}
			goType := CorrectType(baseType)
			subBuilder.WriteString(fmt.Sprintf("%sthis.Args%d = ConvertUtils.GetList<%s>(data);\n", indent3, i, goType))
		} else {
			baseType := CorrectType(t)
			subBuilder.WriteString(fmt.Sprintf("%sthis.Args%d = ConvertUtils.Get<%s>(data[%d]);\n", indent3, i, baseType, i))
		}
	}
	subBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))
	// GetId 方法
	subBuilder.WriteString("\n")
	subBuilder.WriteString(fmt.Sprintf("%spublic int GetId()\n", indent2))
	subBuilder.WriteString(fmt.Sprintf("%s{\n", indent2))
	subBuilder.WriteString(fmt.Sprintf("%svar idProperty = this.GetType().GetProperty(\"ID\");\n", indent3))
	subBuilder.WriteString(fmt.Sprintf("%sif (idProperty != null)\n", indent3))
	subBuilder.WriteString(fmt.Sprintf("%s{\n", indent3))
	subBuilder.WriteString(fmt.Sprintf("%s\treturn (int)idProperty.GetValue(this);\n", indent3))
	subBuilder.WriteString(fmt.Sprintf("%s}\n", indent3))
	subBuilder.WriteString(fmt.Sprintf("%sthrow new Exception($\"当前类 {{this.GetType().Name}}  未定义 ID 属性\");\n", indent3))
	subBuilder.WriteString(fmt.Sprintf("%s}\n", indent2))
	subBuilder.WriteString(fmt.Sprintf("%s}\n", indent1))
	subBuilder.WriteString("\n")
}

// CorrectType 将 Excel 字段类型转为 C# 类型
func CorrectType(fieldType string) string {
	switch strings.ToLower(fieldType) {
	case "int":
		return "int"
	case "float":
		return "float"
	case "double":
		return "double"
	case "string":
		return "string"
	case "bool":
		return "bool"
	case "long":
		return "long"
	case "datetime":
		return "DateTime"
	case "intslice":
		return "List<int>"
	case "boolslice":
		return "List<bool>"
	case "floatslice":
		return "List<float>"
	case "doubleslice":
		return "List<double>"
	case "stringslice":
		return "List<string>"
	case "longslice":
		return "List<long>"
	case "datetimeslice":
		return "List<DateTime>"
	default:
		return "string"
	}
}

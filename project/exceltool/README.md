# Excel2Csharp

本项目用于将 Excel 文件批量转换为 CSV 和 c#文件。

## 依赖安装

如遇依赖下载缓慢或失败，可切换国内 Go 代理：

**阿里云代理：**
```sh
go env -w GOPROXY=https://mirrors.aliyun.com/goproxy/
```

**七牛云代理：**
```sh
go env -w GOPROXY=https://goproxy.cn,direct
```

**恢复官方代理：**
```sh
go env -w GOPROXY=https://proxy.golang.org,direct
```

**修正依赖：**
```sh
go mod tidy
```

## 使用方法

1. 将待转换的 Excel 文件放入 `asset/excel` 目录。
2. 运行主程序：
   ```sh
   go run main.go
   ```
3. 转换后的 CSV 文件会输出到 `asset/csv` 目录。

---
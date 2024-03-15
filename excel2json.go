package excel2json

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"path"
	"strconv"
	"strings"
	"time"
)

// excel 读取excel数据表
//
//	@Description:
func excel() {
	start := time.Now()
	file, jsonDir, header, key, sheet := cmd()
	fmt.Println(file, jsonDir, header, key, sheet)
	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err, file, jsonDir, header, key, sheet)
		return
	}

	rows := xlsx.GetRows(sheet)
	var name []string
	list := make([]map[string]interface{}, 0)
	for i, row := range rows {
		hang := make(map[string]interface{}, 0)

		for num, text := range row {
			if i == key-1 {
				name = append(name, text)
			} else if i > header-1 && name[num] != "" {
				hang[name[num]] = text
				// 转换为int类型，如果可以的话
				int, err := strconv.Atoi(text)
				if err == nil {
					hang[name[num]] = int
				}
				// 转换为float64类型，如果可以的话
				float, err := strconv.ParseFloat(text, 64)
				if err == nil {
					hang[name[num]] = float
				}
			}
		}

		if i >= header {
			list = append(list, hang)
		}
	}

	// 将数据格式化为JSON字符串
	data, err := json.Marshal(list)
	if err != nil {
		fmt.Printf("json.marshal failed,err:%v", err)
		return
	}

	// 如果未指定JSON文件输出路径，则默认与Excel文件同名但扩展名为.json
	if jsonDir == "" {
		fileSuffix := path.Ext(file)                         // 获取文件后缀
		filenameOnly := strings.TrimSuffix(file, fileSuffix) // 获取文件名
		jsonDir = filenameOnly + ".json"
	}
	// 写入JSON文件
	write(jsonDir, string(data))

	// 输出处理耗时
	cost := time.Since(start)
	fmt.Println(file, "is success!", cost)
}

// write函数用于将数据写入指定路径的JSON文件
// @file 文件路径与文件名
// @data 文件内容
func write(file string, data string) {
	obj, err := os.Create(file)
	if err != nil {
		fmt.Println(err)
	}
	obj.WriteString(data)
	obj.Close()
}

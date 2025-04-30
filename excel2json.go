package excel2json

import (
	"encoding/json"
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
	"math"
	"os"
	"path"
	"strings"
	"time"
)

// Excel 将Excel文件转换为JSON格式
//
// @Description: 此函数用于读取指定的Excel文件，根据提供的表头位置、关键行索引和工作表名称，将数据转换为JSON格式，并保存到指定的目录下。
// @param file string - Excel文件的路径。
// @param jsonDir string - 生成的JSON文件将保存在该目录下。如果未指定，将与Excel文件同名但扩展名为.json。
// @param header int - 表头所在的行索引（从0开始）。
// @param key int - 包含列名的关键行索引（从0开始）。
// @param sheet string - Excel工作表的名称。
func Excel(file string, jsonDir string, header int, key int, sheet string) {
	// 记录函数开始时间
	start := time.Now()
	// 打印函数参数，用于调试
	fmt.Println(file, jsonDir, header, key, sheet)

	// 使用excelize库打开Excel文件
	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		// 打印错误信息并返回
		fmt.Println(err, file, jsonDir, header, key, sheet)
		return
	}

	// 获取指定工作表的所有行
	rows, err := xlsx.GetRows(sheet)
	var name []string                         // 用于存储列名
	list := make([]map[string]interface{}, 0) // 用于存储转换后的数据

	// 遍历所有行，根据列名和数据类型，构建数据列表
	for i, row := range rows {
		hang := make(map[string]interface{})

		for num, text := range row {
			if i == key-1 {
				name = append(name, text)
			} else if i > header-1 && len(name) > num {
				//去掉字符串前后空格
				text = strings.TrimSpace(text)
				hang[name[num]] = text

				// 尝试将字符串转换为bool类型
				boolData, err2 := strconv.ParseBool(text)
				if err2 == nil {
					hang[name[num]] = boolData
					continue
				}

				// 尝试将字符串转换为float64类型
				floatData, err2 := strconv.ParseFloat(text, 64)
				if err2 == nil {
					hang[name[num]] = sanitizeValue(floatData)
					continue
				}

				// 尝试将字符串转换为int类型
				intData, err2 := strconv.Atoi(text)
				if err2 == nil {
					hang[name[num]] = intData
					continue
				}

			}
		}

		if i >= header {
			list = append(list, hang)
		}
	}

	// 将数据列表转换为JSON字符串
	data, err := json.MarshalIndent(list, "", "\t")
	if err != nil {
		// 转换失败，打印错误信息并返回
		fmt.Printf("json.marshal failed,err:%v,内容如下:%s", err, data)
		return
	}

	// 处理JSON文件输出路径，未指定时默认与Excel文件同名但扩展名为.json
	if jsonDir == "" {
		fileSuffix := path.Ext(file)                         // 获取Excel文件的后缀名
		filenameOnly := strings.TrimSuffix(file, fileSuffix) // 获取不带后缀的文件名
		jsonDir = filenameOnly + ".json"                     // 构造JSON文件名
	}
	// 将JSON数据写入文件
	write(jsonDir, string(data))

	// 打印处理耗时
	cost := time.Since(start)
	fmt.Println(file, "is success!", cost)
}

// write 函数用于将数据写入指定路径的JSON文件
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

// 解析json不支持的正无穷大
func sanitizeValue(v float64) interface{} {
	if math.IsInf(v, 1) {
		return "Infinity"
	} else if math.IsInf(v, -1) {
		return "-Infinity"
	} else if math.IsNaN(v) {
		return "NaN"
	}
	return v
}

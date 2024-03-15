package excel2json

import (
	"flag"
)

func main() {
	excel()
}

// excel 读取excel数据表
//
//	@returns 无返回值
func excel() {
	file, jsonDir, header, key, sheet := cmd() // 通过cmd函数获取必要的参数
	Excel(file, jsonDir, header, key, sheet)   // 调用Excel函数进行数据处理和转换
}

// cmd函数用于解析命令行传入的参数，并返回相关配置
// 参数：
// - excelDir: 指定输入的Excel文件路径的字符串指针
// - jsonDir: 指定输出的json文件路径的字符串指针
// - header: 表示表格中表头占用的行数的整数指针
// - key: 表示key值所在的行数的整数指针
// - sheet: 指定excel表的页签名称的字符串指针
// 返回值：
// - excelDir: 输入的Excel文件路径字符串
// - jsonDir: 指定的json文件输出路径字符串
// - header: 表头占用的行数整数
// - key: key值所在的行数整数
// - sheet: excel表的页签名称字符串
func cmd() (excelDir string, jsonDir string, header int, key int, sheet string) {
	excelDir2 := flag.String("e", "", "Required. 输入的Excel文件路径")
	jsonDir2 := flag.String("j", "", "指定输出的json文件路径")
	header2 := flag.Int("h", 3, "表格中有几行是表头.")
	key2 := flag.Int("k", 2, "key值在第几行")
	sheet2 := flag.String("s", "data", "excel表的页签名称")
	flag.Parse() // 解析命令行参数

	excelDir = *excelDir2
	jsonDir = *jsonDir2
	header = *header2
	key = *key2
	sheet = *sheet2
	return
}

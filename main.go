package main;

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"path"
	"strconv"
	"strings"
	"time"
)

func main() {
	excel();
}

func cmd() (string, string, int, int, string) {
	excel_dir := flag.String("e", "", "Required. 输入的Excel文件路径")
	json_dir := flag.String("j", "", "指定输出的json文件路径");
	header := flag.Int("h", 3, "表格中有几行是表头.");
	key := flag.Int("k", 2, "key值在第几行");
	sheet := flag.String("s", "data", "excel表的页签名称")
	flag.Parse() //解析输入的参数
	return *excel_dir, *json_dir, *header, *key, *sheet;
}

/**
	读取excel数据表
 */
func excel() {
	start := time.Now();
	file, json_dir, header, key,sheet  := cmd();
	xlsx, err := excelize.OpenFile(file);
	//xlFile, err := xlsx.OpenFile(file);
	if err != nil {
		fmt.Println(err)
		return
	}

	rows := xlsx.GetRows(sheet);
	var name [] string;
	list := make([]map[string]interface{}, 0);
	for i, row := range rows {
		hang := make(map[string]interface{}, 0);

		for num, text := range row {
			if i == key-1 {
				//fmt.Println(text);
				name = append(name, text);
			} else if i > header-1 && name[num] != "" {

				hang[name[num]] = text;
				//判断并转为int
				int, err := strconv.Atoi(text);
				if err == nil {
					hang[name[num]] = int;
				}
				//判断并转为float64
				float, err := strconv.ParseFloat(text, 64);
				if err == nil {
					hang[name[num]] = float;
				}

				//fmt.Printf("%s\n", text)
			}

		}

		if i >= header {
			list = append(list, hang);
		}

	}

	//格式化为json
	data, err := json.Marshal(list);
	//fmt.Println(string(data));
	//data, err := json.MarshalIndent(list, "", "    ");

	if err != nil {
		fmt.Printf("json.marshal failed,err:", err)
		return
	}
	//fmt.Println(string(data));

	if json_dir == "" {
		fileSuffix := path.Ext(file)                         //获取文件后缀
		filenameOnly := strings.TrimSuffix(file, fileSuffix) //获取文件名
		json_dir = filenameOnly + ".json";
	}
	write(json_dir, string(data));

	//运行时间
	cost := time.Since(start);
	fmt.Println(file, "is success!", cost);
}

/**
	写入json文件
	@ file 文件路径与文件名
	@ data 文件内容
 */
func write(file string, data string) {

	obj, err := os.Create(file);
	if err != nil {
		fmt.Println(err);
	}
	obj.WriteString(data);
	obj.Close();
}

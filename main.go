package main;

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"path"
	"strings"
)
func main(){

	excel();

}

func cmd()(string,string,int,int,string){
	excel_dir := flag.String("e", "", "Required. 输入的Excel文件路径")
	json_dir := flag.String("j", "", "指定输出的json文件路径");
	header := flag.Int("h", 3, "表格中有几行是表头.");
	key := flag.Int("k", 1, "key值在第几行");
	sheet := flag.String("s","data","excel表的页签名称")

	flag.Parse() //解析输入的参数
	return *excel_dir,*json_dir,*header,*key,*sheet;

}

/**
	读取excel数据表
 */
func excel(){

	file,json_dir,header,key,sheet := cmd();

	xlsx, err := excelize.OpenFile(file);
	if err != nil {
		fmt.Println(err)
		return
	}
	rows := xlsx.GetRows(sheet);
	var lie[] string;
	list := make(map[int]map[string]string);

	for i, row := range rows {
		hang := make(map[string]string);
		//从第三行开始读取正文
		if i >= header {
			for num, colCell := range row {
				if lie[num] != "" {
					//heng = append(heng,colCell);
					hang[lie[num]] = colCell;
				}
			}
			//排除前面空余的行
			list[i - header] = hang;
			//fmt.Println(hang);
		}else if i+1 == key{		//获取第二行的key值
			for _, colCell := range row {
				lie = append(lie,colCell);
			}
		}


	}


	//格式化为json
	data, err := json.Marshal(list);
	if err != nil {
		fmt.Printf("json.marshal failed,err:", err)
		return
	}

	//fmt.Println(string(data));
	//fmt.Println(data);

	if json_dir == "" {
		fileSuffix := path.Ext(file) //获取文件后缀
		filenameOnly := strings.TrimSuffix(file, fileSuffix)//获取文件名
		json_dir = filenameOnly+".json";
	}
	write(json_dir,string(data));

}

/**
	写入json文件
	@ file 文件路径与文件名
	@ data 文件内容
 */
func write(file string,data string){

	obj, err := os.Create(file);
	if err != nil {
		fmt.Println(err);
	}
	obj.WriteString(data);
	obj.Close();

}



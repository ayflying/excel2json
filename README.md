# go语言版的excel2json
### go语言的高性能，而且支持跨平台
go语言版本的excel2json，可以把excel表转换成json文件

使用方法：
在命令行中使用，输入
```bash
cxcel2json -help
```
运行结果，可查询运行参数
```bash
  -e string
        Required. 输入的Excel文件路径
  -h int
        表格中有几行是表头. (default 3)
  -j string
        指定输出的json文件路径
  -k int
        key值在第几行 (default 1)
  -s string
        excel表的页签名称 (default "data")
```
如果需要转一个excel表比如 item.xlxs 表
```bash
excel2json -e item.xlxs -j item.json
```

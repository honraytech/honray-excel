# honray-excel
thinkphp5 系统生成excel函数

## 安装


> composer require honray/tp5-excel


##使用

###在后端生成excel,并返回路径


#####执行方法export_excel($data, $title, $version = '2007', $savefile = '')



* 在服务器生成excel表并-返回路径
* 
* -------------------------------------------------------------------------
* @param string $data 需要导出的数据，必须
* @param array() $title 导出的excel表的表头信息，必须
* @param string $version 导出的excel信息，默认为2007，可选
* @param string $savefile 导出的excel保存的名字，可选(需要英文，字符串)
* @param array() $width 导出的excel保存的宽度，必选，最好填写,具体到每列的宽度如$width = array('A'=>10,'B'=>10,'C'=>25,'D'=>20,'E'=>25,'F'=>25,'G'=>25,'H'=>25);
* @param array(fontfamily,size,color,bold) $font 导出的excel表头的样式,可选(已废弃，损耗性能)
* @param string $sheetname 导出的excel表的名字，默认为“shee1”，可选
* 返回值:文件名称$url:如lession.xls
* -------------------------------------------------------------------
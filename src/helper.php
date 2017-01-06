<?php
    
/**
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
**/
use think\excel\PHPExcel;

function export_excel($data, $title, $version = '2007', $savefile = '', $width = '', $font= ['fontfamily'=>'SimHei','size'=>'12','color'=>'000000','bold'=>'false'])
{
    ini_set('max_execution_time', '0');
    //若没有指定文件名
    if (empty($savefile)) {
       $savefile='data';
    }

    //若没有指定宽度
    if (empty($width)) {
       $width = ['A' => 25,'B' => 25,'C' => 25,'D' => 25,'E' => 25,'F' => 25,'G' => 25,'H' => 25];
    }

    //判断版本输出和后缀名
    if ($version=='2007') {
        $ext='.xlsx';
    }
    elseif ($version=='2005') {
        $ext='.xls';
    }
    else {
        return '输入的版本参数有误';
    }

    //若指字了excel表头，则把表单追加到正文内容前面去
    if (is_array($title)) {
        array_unshift($data, $title);
    }

    //实例化excel
    $objPHPExcel = new PHPExcel();

    //表示用第$i张表格         
    $obj = $objPHPExcel->setActiveSheetIndex(0);

    foreach ($data as $k => $v) {                                   
        //从第二行插入数据；
        $row = $k+1; //行
        $nn = 0;//列                     
        
        foreach ($v as $kk=>$vv) {
            if($nn<26){
                 $col = chr(65 + $nn); //列，其中chr()是把数据转换为ASCII码
                 
                 $obj->setCellValue($col . $row, $vv); //列,行,值

                $obj->getStyle($col.$row)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //水平居中
                 $nn++;
            }
            else{
             $col1= chr(65+$nn-26); //列，其中chr()是把数据转换为ASCII码
             $col=A.$col1;
             $obj->setCellValue($col.$row, $vv); //列,行,值

            $obj->getStyle($col.$row)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //水平居中
             $nn++;
            }
        }
    }

    foreach($width as $k => $v){
        $objPHPExcel->getActiveSheet()->getColumnDimension($k)->setWidth($v);
    }
    
    $objPHPExcel->getActiveSheet()->setTitle($sheetname); //题目  

    //保存到服务器
    $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $filename = $savefile.$ext;

    if (!is_dir('./uploads/excel/')) {
        @mkdir('./uploads/excel/');
        @chmod('./uploads/qrcode/', 0777);
    }

    $objWriter->save('uploads/excel/'.$filename);
    return 'uploads/excel/'.$filename;  
}

?>
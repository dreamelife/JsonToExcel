<?php
include_once('./Classes/PHPExcel.php');
include_once('./Classes/Base.php');
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

function dump($data){
    echo '<pre>';
    print_r($data);
    echo '</pre>';
}
$data['headers'] =  [
    'name'=>['value'=>'姓名','attrs'=>['color'=>'FFA07A','rows'=>4,'cols'=>5]],
    'age'=>['value'=>'年龄'],
    'sex'=>['value'=>'性别']
];
$data['rows'] = [
    'row1'=>[
        'name'=>['value'=>'dylan','attrs'=>['color'=>'FF34B3','rows'=>2,'cols'=>1]],
        'age' =>['value'=>'27'],
        'shao' =>['value'=>'bunanbunv'],
    ],
    'row2'=>[
        'name'=>['value'=>'bill','attrs'=>['color'=>'B0C4DE','rows'=>1]],
        'age' =>['value'=>'28'],
        'sex' =>['value'=>'男','attrs'=>['color'=>'9F79EE','rows'=>1,'cols'=>1]],
    ],
    'row3'=>[
        'name'=>['value'=>'bill','attrs'=>['color'=>'4876FF','rows'=>1,'cols'=>1]],
        'age' =>['value'=>'600'],
        'sex' =>['value'=>'男'],
    ]
];

$data1 = $data;

$data1['headers'] =  [
    'name'=>['value'=>'name'],
    'age'=>['value'=>'age'],
    'sex'=>['value'=>'sex']
];
$datas['diyige'] = $data;
$datas['shee2'] = $data1;


//----------------------------------------------------------------------------------------------------------------------------------
$objPHPExcel = new PHPExcel();
$Base = new Base();
$sheet = 0 ;
foreach ($datas as $key=>$data){
    //初始化起始变量
    $pos_row = '0';
    $pos_col = 'A';
    //创建sheet
    if($sheet > 0){
        $objPHPExcel->createSheet($sheet);
        $objPHPExcel->setActiveSheetIndex($sheet);
        $objActSheet = $objPHPExcel->getActiveSheet($sheet);
    }else{
        $objActSheet = $objPHPExcel->getActiveSheet(0);
    }
    $sheet++;
    //------------------数据填充-------------------------------------------------------

    //-----------------set sheet's name-----------
    $objActSheet->setTitle($key);
    //-----------------set header ----------------
    $headers = $data['headers'];
    if($headers){
        //-----------------------------------------
        $header_col = $pos_col;
        $pos_row_changed = $pos_row;
        //-----------------------------------------
        foreach ($headers as $key=>$header){
            //--------------------attr-----------------
            $attrs = isset($header['attrs']) ? $header['attrs'] : '';
            list($attrs,$position,$header_col,$pos_row_new) = $Base->set_position($header_col,$pos_row,$attrs,$objActSheet);
            $pos_row_changed = max($pos_row_changed,$pos_row_new);
            //------------------set attr---------------
            if($attrs){
                foreach ($attrs as $action=>$attr){
                    if(empty($action)) continue;
                    $Base->$action($position,$attr,$objActSheet);
                }
            }
            $objActSheet->setCellValue($position,$header['value']);
            $header_col =  chr(ord($header_col)+1);
        }
        //-----------------------------------------
        $header_col = $pos_col;
        $pos_row = $pos_row_changed;
        //-----------------------------------------
    }

    $rows = $data['rows'];
    if(empty($rows)) continue;
    foreach ($rows as $row){
        $row_col = $pos_col;
        $pos_row_changed = $pos_row;
        //-----------------------------------------
        foreach ($row as $key=>$v){
            $attrs = isset($v['attrs']) ? $v['attrs'] : '';
            list($attrs,$position,$row_col,$pos_row_new) = $Base->set_position($row_col,$pos_row,$attrs,$objActSheet);
            $pos_row_changed = max($pos_row_changed,$pos_row_new);
            if($attrs){
                foreach ($attrs as $action=>$attr){
                    if(empty($action)) continue;
                    $Base->$action($position,$attr,$objActSheet);
                }
            }

            $row_col =  chr(ord($row_col)+1);
            if( strlen($v['value']) == 0 ) continue;
            $objActSheet->setCellValue($position,$v['value']);
        }
        //-----------------------------------------
        $pos_row = $pos_row_changed;
        //-----------------------------------------
    }

}
//exit;
$name = '超级newbie.xlsx';
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
ob_end_clean();//清除缓冲区,避免乱码
// Redirect output to a client’s web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename='.$name.'');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');
// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
$objWriter->save('php://output');
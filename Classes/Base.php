<?php
class Base extends PHPExcel{

    /**
     * @param $position
     * @param $color
     * @param $objActSheet
     * @author liuyang
     */
    public function color($position,$color,$objActSheet){
        $objStyle = $objActSheet->getStyle($position);
        $objFont = $objStyle->getFont();
        $objFont->getColor()->setRGB($color);
    }

    /**
     * @param $header_row
     * @param $pos_col
     * @param $attrs
     * @param $objActSheet
     * @return array
     * @author liuyang
     */
    public function set_position($header_col,$pos_row,$attrs,$objActSheet){
        $position_initial = $header_col.($pos_row + 1);
        $pos_row_initial = $pos_row;

        if(isset($attrs['cols'])){
            $header_col = chr(ord($header_col)+$attrs['cols'] - 1);
            unset($attrs['cols']);
        }
        if(isset($attrs['rows'])){
            $pos_row  = $pos_row + $attrs['rows'];
            unset($attrs['rows']);
        }else{
            $pos_row =  $pos_row + 1;
        }
        if($pos_row == 0){
            $pos_row = 1;
        }
        $position = $header_col.$pos_row;
//        echo '----------------------';
//        dump($pos_row);
//        dump($position_initial);
//        dump($position);
//        dump($attrs);
//        echo '----------------------';
//        dump($position);
        if($position_initial !== $position){
            //合并单元格
            //dump($position_initial);
            //dump($position);
            $objActSheet->mergeCells("$position_initial:$position");
        }
        $data[] = $attrs;
        $data[] = $position_initial;
        $data[] = $header_col;
        $data[] = $pos_row;
        return $data;
    }

    public function size(){
//        $objStyle = $objActSheet->getStyle($position);
//        $objAlign = $objStyle->getAlignment();
//        $objAlign->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);  //上下居中
//        //字体及颜色
//        $objFont = $objStyle->getFont();
//        $objFont->setName('黑体');
//
//        $objFont->getColor()->setARGB($color);
    }


}

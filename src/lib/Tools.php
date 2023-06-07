<?php

namespace Byk\Excel\lib;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Tools
{

    /**
     * 数字转为字母
     * @param int $num 数字
     * @return string
     */
    public static function numToLetter(int $num)
    {
        if ($num<=0){
            return '';
        }
        $num = $num - 1;
        //获取A的ascii码
        $ordA = ord('A');
        return chr($ordA + $num);
    }

    /**
     * 数字转为excel表格字母列
     * @param int $num 数字
     * @return string
     */
    public static function numToExcelLetter(int $num)
    {
        //由于大写字母只有26个，所以基数为26
        $base = 26;
        $result = '';
        while ($num > 0 ) {
            $mod = (int)($num % $base);
            $num = (int)($num / $base);
            if($mod == 0){
                $num -= 1;
                $temp = self::numToLetter($base) . $result;
            } elseif ($num == 0) {
                $temp = self::numToLetter($mod) . $result;
            } else {
                $temp = self::numToLetter($mod) . $result;
            }
            $result = $temp;
        }
        return $result;
    }

    public static function processingData(Worksheet $activeSheet,array $headers,array $data)
    {
        foreach ($headers as $header)
        {
            [
                'field' => $field,
                'title' => $title,
                'wrap_text' => $wrapText,
                'width' => $width,
                'letter' => $letter,
            ] = $header;
            $activeSheet->setCellValue($letter . '1', $title);
            if ($wrapText) {
                $activeSheet->getStyle($letter)->getAlignment()->setWrapText(true);
            }
            if ((int)$width > 0) {
                $activeSheet->getColumnDimension($letter)->setWidth($width);
            }
            $row = 2;
            foreach ($data as $value) {
                $content = '';
                if (!empty($field)) {
                    $content = $value[$field] ?? $field;
                }
                $activeSheet->setCellValue($letter . $row, $content);
                $row++;
            }
        }
        return $activeSheet;
    }

    public static function processingData2(Worksheet $activeSheet,array $headers,array $data)
    {
        foreach ($headers as $header){
            [
                'field' => $field,
                'title' => $title,
                'merge' => $merge,
                'letter' => $letter,
            ] = $header;
            $activeSheet->setCellValue($letter.'1',$title);
            $row = 2;
            foreach ($data as $item){
                $startRow = $row;
                foreach ($item as $value){
                    $content = '';
                    if (!empty($field) and isset($value[$field])){
                        $content = $value[$field];
                    }
                    $activeSheet->setCellValue($letter.$row,$content);
                    $row++;
                }
                $endRow = $row-1;
                if ($merge===true and $endRow>$startRow){
                    // 合并单元格
                    $activeSheet->mergeCells($letter.$startRow.':'.$letter.$endRow);
                }
            }
        }
    }
}
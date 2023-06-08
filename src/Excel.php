<?php

namespace Byk\Excel;

use Byk\Excel\lib\Tools;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

/**
 *
 */
class Excel
{
    /**
     * @var object
     */
    private static $instance;
    /**
     * @var array
     */
    private $header = [];
    /**
     * @var array
     */
    private $data = [];
    /**
     * @var int
     */
    private $dataType = 1;
    /**
     * @var int
     */
    private $width = 0;
    /**
     * @var object
     */
    private $excel;

    private $fileType = 'xlsx';

    /**
     * 初始化
     * @return static
     */
    public static function instance()
    {
        if (is_null(self::$instance)){
            self::$instance = new static();
        }
        return self::$instance;
    }

    /**
     * 设置标题
     * @param array $header
     * @return $this
     */
    public function setHeader(array $header)
    {
        if(count($header)==count($header,1)){
            throw new \Exception('标题数据必须是二维数组');
        }
        $num = 1;
        foreach ($header as &$value){
            if (!isset($value['headers'])){
                if (!isset($value['title'])){
                    throw new \Exception('缺失标题title参数'.key($value));
                }
                $value['letter'] = Tools::numToExcelLetter($num);
                $value = array_merge(
                    [
                        'field' => '',
                        'wrap_text' => false,
                        'width' => 0,
                        'merge'=>false
                    ],
                    $value
                );
                $num++;
            }else{
                $num = 1;
                foreach ($value['headers'] as &$item){
                    if (!isset($item['title'])){
                        throw new \Exception('缺失标题title参数');
                    }
                    $item['letter'] = Tools::numToExcelLetter($num);
                    $item = array_merge(
                        [
                            'field' => '',
                            'wrap_text' => false,
                            'width' => 0,
                            'merge'=>false
                        ],
                        $item
                    );
                    $num++;
                }
            }
        }
        unset($value);
        $this->header = $header;
        return $this;
    }

    /**
     * 设置数据
     * @param array $data
     * @return $this
     */
    public function setData(array $data)
    {
        $this->data = $data;
        return $this;
    }

    /**
     * 设置数据类型
     * @param int $dataType 数据类型，1.二维数组，2.三维数组 3.多个工作表
     * @return $this
     */
    public function setDataType(int $dataType = 1)
    {
        $this->dataType = $dataType;
        return $this;
    }

    public function setFileType(string $fileType = 'xlsx')
    {
        $this->fileType = $fileType;
        return $this;
    }

    /**
     * 设置单元格默认宽度
     * @param int $width
     * @return $this
     */
    public function setWidth(int $width)
    {
        $this->width = $width;
        return $this;
    }

    /**
     * 生成数据
     * @return $this
     */
    public function create($callback = null)
    {
        $this->excel = new Spreadsheet();
        $this->excel->setActiveSheetIndex(0);
        $activeSheet = $this->excel->getActiveSheet();
        if ($this->width>0){
            $activeSheet->getDefaultColumnDimension()->setWidth($this->width); //设置列默认宽度
        }
        if (is_callable($callback)) {
            call_user_func_array($callback, [$this->excel,$this->header, $this->data]);
        } elseif ($this->dataType == 2) {
            Tools::processingData2($activeSheet, $this->header, $this->data);
        } elseif ($this->dataType == 3){
            Tools::processingSheetData($this->excel,$this->header,$this->data,['width'=>$this->width]);
        } else {
            Tools::processingData($activeSheet, $this->header, $this->data);
        }
        return $this;
    }

    /**
     * 保存
     * @param string $fileName
     * @return string|true
     */
    public function save(string $fileName)
    {
        try {
            $writer = IOFactory::createWriter($this->excel, ucfirst($this->fileType));
            $writer->save($fileName);
            return true;
        } catch (Exception $e) {
            return $e->getMessage();
        }
    }

    public function download($fileName)
    {
        if ($this->fileType=='xlsx'){
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        }else{
            header('Content-Type: application/vnd.ms-excel');
        }
        header("Content-Disposition: attachment;filename="
            . $fileName);
        header('Cache-Control: max-age=0');
        $this->save('php://output');
        exit();
    }
}
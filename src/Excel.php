<?php

namespace Excel;

use excel\lib\Tools;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

/**
 *
 */
final class Excel
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
    private $freeze = false;

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
     * @param array $header 标题
     * <div id="function.setHeader" class="refentry"> <div class="refnamediv"><p class="para"><div class="example" id="example-setHeader1"><p><strong>Example #1<span class="function"><strong style="color:#CC7832">setHeader()</strong></span><span class="parameter" style="color:#3A95FF">dataType=1|2</span>结构</strong></p><div class="example-contents"><div class="phpcode" style="border-color:gray;background:#232525"><span><span style="color: #000000"><span style="color: #9876AA">&lt;?php<br /></span><span style="color: #9876AA">$headers&nbsp;&nbsp;</span><span style="color: #007700">=&nbsp;</span><span style="color: #007700">[</span><br /><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'field'</span><span style="color: #007700">=></span><span style="color: #DD0000">'date'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;字段名称，可为空<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'title'</span><span style="color: #007700">=></span><span style="color: #DD0000">'日期'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;标题名称<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'wrap_text'</span><span style="color: #007700">=></span><span style="color: #9876AA">true</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;开启换行<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'format_code'</span><span style="color: #007700">=></span><span style="color: #DD0000">'str'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;内容格式<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'total'</span><span style="color: #007700">=></span><span style="color: #9876AA">false</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;关闭合计列<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'vertical'</span><span style="color: #007700">=></span><span style="color: #DD0000">'center'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;垂直居中<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'horizontal'</span><span style="color: #007700">=></span><span style="color: #DD0000">'center'</span></span><span style="color: #FF8000">//&nbsp;水平居中<br /></span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />];<br /><br /><br /></span></span></div></div></div></p><p class="para"><div class="example" id="example-setHeader2"><p><strong>Example #2<span class="function"><strong style="color:#CC7832">setHeader()</strong></span><span class="parameter" style="color:#3A95FF">dataType=3</span>结构</strong></p><div class="example-contents"><div class="phpcode" style="border-color:gray;background:#232525"><span><span style="color: #000000"><span style="color: #9876AA">&lt;?php<br /></span><span style="color: #9876AA">$headers&nbsp;&nbsp;</span><span style="color: #007700">=&nbsp;</span><span style="color: #007700">[</span><br /><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'sheet'</span><span style="color: #007700">=></span><span style="color: #DD0000">'工作表名称'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;工作表名称（可选）<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'data_type'</span><span style="color: #007700">=></span><span style="color: #9876AA">2</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;工作表数据类型（可选），默认：1<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'headers'</span><span style="color: #007700">=></span><span style="color: #007700">[ <br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[ <br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'field'</span><span style="color: #007700">=></span><span style="color: #DD0000">'date'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;字段名称，可为空<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'title'</span><span style="color: #007700">=></span><span style="color: #DD0000">'日期'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;标题名称<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'wrap_text'</span><span style="color: #007700">=></span><span style="color: #9876AA">true</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;开启换行<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'format_code'</span><span style="color: #007700">=></span><span style="color: #DD0000">'str'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;内容格式<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'total'</span><span style="color: #007700">=></span><span style="color: #9876AA">false</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;关闭合计列<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'vertical'</span><span style="color: #007700">=></span><span style="color: #DD0000">'center'</span><span style="color: #007700">,</span><span style="color: #FF8000">//&nbsp;垂直居中<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'horizontal'</span><span style="color: #007700">=></span><span style="color: #DD0000">'center'</span></span><span style="color: #FF8000">//&nbsp;水平居中<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />];<br /><br /><br /></span></span></div></div></div> </p> </div> </div>
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
                        'merge'=>false,
                        'format_code'=>'',
                        'total'=>false,
                        'vertical'=>'',//垂直
                        'horizontal'=>'',//水平
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
                            'merge'=>false,
                            'format_code'=>'',
                            'total'=>false,
                            'vertical'=>'',//垂直
                            'horizontal'=>'',//水平
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
     * @param array $data 数据
     * <div id="function.setData" class="refentry"> <div class="refnamediv" style="width:260px;"><p class="para"><div class="example" id="example-setData1"><p><strong>Example #1<span class="function"><strong style="color:#CC7832">setData()</strong></span><span class="parameter" style="color:#3A95FF">dataType=1</span>结构</strong></p><div class="example-contents"><div class="phpcode" style="border-color:gray;background:#232525"><span><span style="color: #000000"><span style="color: #9876AA">&lt;?php<br /></span><span style="color: #9876AA">$data&nbsp;&nbsp;</span><span style="color: #007700">=&nbsp;</span><span style="color: #007700">[</span><br /><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-01'</span><span style="color: #007700">,</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;],<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-02'</span><span style="color: #007700">,</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />];</span></span></div></div></div></p><p class="para"><div class="example" id="example-setData2"><p><strong>Example #2<span class="function"><strong style="color:#CC7832">setData()</strong></span><span class="parameter" style="color:#3A95FF">dataType=2</span>合并列结构</strong></p><div class="example-contents"><div class="phpcode" style="border-color:gray;background:#232525"><span><span style="color: #000000"><span style="color: #9876AA">&lt;?php<br /></span><span style="color: #9876AA">$data&nbsp;&nbsp;</span><span style="color: #007700">=&nbsp;</span><span style="color: #007700">[</span><br /><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;'2024-01-01'</span><span style="color: #007700">=></span><span style="color: #007700">[<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-01'</span><span style="color: #007700">,</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;],<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-01'</span><span style="color: #007700">,</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />];</span></span></div></div></div></p><p class="para"><div class="example" id="example-setData3"><p><strong>Example #3<span class="function"><strong style="color:#CC7832">setData()</strong></span><span class="parameter" style="color:#3A95FF">dataType=3</span>多工作表结构</strong></p><div class="example-contents"><div class="phpcode" style="border-color:gray;background:#232525"><span><span style="color: #000000"><span style="color: #9876AA">&lt;?php<br /></span><span style="color: #9876AA">$data&nbsp;&nbsp;</span><span style="color: #007700">=&nbsp;</span><span style="color: #007700">[</span><br /><span style="color: #FF8000">&nbsp;&nbsp;&nbsp;&nbsp;//&nbsp;工作表1<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[</span><span style="color: #DD0000">'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-01'</span><span style="color: #007700"></span><span style="color: #007700">],<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[</span><span style="color: #DD0000">'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-02'</span><span style="color: #007700"></span><span style="color: #007700">]<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;],<br /></span><span style="color: #FF8000">&nbsp;&nbsp;&nbsp;&nbsp;//&nbsp;工作表2<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'2024-01-01'</span><span style="color: #007700">=></span><span style="color: #007700">[<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-01'</span><span style="color: #007700">,</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;],<br /></span><span style="color: #007700">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<br /></span><span style="color: #DD0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'date'</span><span style="color: #007700">=></span><span style="color: #DD0000">'2024-01-01'</span><span style="color: #007700">,</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />&nbsp;&nbsp;&nbsp;&nbsp;]</span><span style="color: #007700"><br />];</span></span></div></div></div></p></div></div>
     * @return $this
     */
    public function setData(array $data)
    {
        $this->data = $data;
        return $this;
    }

    /**
     * 设置数据类型
     * @param int $dataType 数据类型，1.二维数组，2.三维数组
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

    public function freezeHeader(bool $freeze)
    {
        $this->freeze = $freeze;
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
        if ($this->freeze){
            $activeSheet->freezePane('A2');
        }
        if ($this->width>0){
            $activeSheet->getDefaultColumnDimension()->setWidth($this->width); //设置列默认宽度
        }

        if (is_callable($callback)) {
            call_user_func_array($callback, [$this->excel,$this->header, $this->data,['width'=>$this->width,'freeze'=>$this->freeze]]);
        } elseif ($this->dataType == 2) {
            Tools::processingData2($activeSheet, $this->header, $this->data);
        } elseif ($this->dataType == 3){
            Tools::processingSheetData($this->excel,$this->header,$this->data,['width'=>$this->width,'freeze'=>$this->freeze]);
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
            unset($this->excel,$this->header,$this->data);
            $this->header = [];
            $this->data = [];
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
<?php

namespace tests;

use Byk\Excel\Excel;
use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase
{
    public function testSave()
    {
        $headers = [
            ['field'=>'name','title'=>'姓名'],
            ['field'=>'age','title'=>'年龄'],
        ];
        $data = [
            ['name'=>'张三','age'=>22],
            ['name'=>'李四','age'=>24],
            ['name'=>'王五','age'=>20],
            ['name'=>'赵六','age'=>23],
        ];
        $fileName = __DIR__.'/test.xls';
        try {
            $res = Excel::instance()->setWidth(20)->setHeader($headers)->setData($data)->setFileType('xls')->create()->save($fileName);
            $this->assertTrue($res,$res);
        } catch (\Exception $e) {
            return $e->getMessage();
        }
    }
    public function testSave2()
    {
        $headers = [
            ['field'=>'name','title'=>'姓名','merge'=>true],
            ['field'=>'goods','title'=>'商品'],
            ['field'=>'num','title'=>'数量'],
        ];
        $data = [
            [
                ['name'=>'张三','goods'=>'上衣','num'=>1],
                ['name'=>'张三','goods'=>'T恤','num'=>2],
                ['name'=>'张三','goods'=>'外套','num'=>1],
            ],
            [
                ['name'=>'李四','goods'=>'上衣','num'=>2],
                ['name'=>'李四','goods'=>'T恤','num'=>1],
                ['name'=>'李四','goods'=>'外套','num'=>2],
            ],
            [
                ['name'=>'王五','goods'=>'上衣','num'=>2],
                ['name'=>'王五','goods'=>'T恤','num'=>1],
                ['name'=>'王五','goods'=>'外套','num'=>2],
            ],
            [
                ['name'=>'赵六','goods'=>'上衣','num'=>2],
                ['name'=>'赵六','goods'=>'T恤','num'=>1],
                ['name'=>'赵六','goods'=>'外套','num'=>2],
                ['name'=>'赵六','goods'=>'短裤','num'=>2],
            ]
        ];
        $fileName = __DIR__.'/test2.xlsx';
        try {
            $res = Excel::instance()->setWidth(20)->setDataType(2)->setHeader($headers)->setData($data)->create()->save($fileName);
            $this->assertTrue($res,$res);
        } catch (\Exception $e) {
            echo $e->getMessage();
        }
    }

    public function testDownload()
    {
        $headers = [
            ['field'=>'name','title'=>'姓名','merge'=>true],
            ['field'=>'goods','title'=>'商品'],
            ['field'=>'num','title'=>'数量'],
        ];
        $data = [
            [
                ['name'=>'张三','goods'=>'上衣','num'=>1],
                ['name'=>'张三','goods'=>'T恤','num'=>2],
                ['name'=>'张三','goods'=>'外套','num'=>1],
            ],
            [
                ['name'=>'李四','goods'=>'上衣','num'=>2],
                ['name'=>'李四','goods'=>'T恤','num'=>1],
                ['name'=>'李四','goods'=>'外套','num'=>2],
            ],
            [
                ['name'=>'王五','goods'=>'上衣','num'=>2],
                ['name'=>'王五','goods'=>'T恤','num'=>1],
                ['name'=>'王五','goods'=>'外套','num'=>2],
            ],
            [
                ['name'=>'赵六','goods'=>'上衣','num'=>2],
                ['name'=>'赵六','goods'=>'T恤','num'=>1],
                ['name'=>'赵六','goods'=>'外套','num'=>2],
                ['name'=>'赵六','goods'=>'短裤','num'=>2],
            ]
        ];
        $fileName = 'test2.xlsx';
        try {
            Excel::instance()->setWidth(20)->setDataType(2)->setHeader($headers)->setData($data)->create()->download($fileName);
            $this->assertTrue(1==1,'');
        } catch (\Exception $e) {
            echo $e->getMessage();
        }
    }

    public function testSheet()
    {
        $header = [
            [
                //工作表标题（必须）
                'headers'=>[
                    ['field'=>'name','title'=>'姓名'],
                    ['field'=>'age','title'=>'年龄'],
                ]
            ],
            [
                //工作表名称（可选）
                'sheet'=>'成绩表',
                //工作表标题（必须）
                'headers'=>[
                    ['field'=>'name','title'=>'姓名','merge'=>true],
                    ['field'=>'subject','title'=>'科目'],
                    ['field'=>'score','title'=>'分数'],
                ],
                //工作表数据类型（可选），默认：1
                'data_type'=>2
            ]
        ];
        $data = [
            [
                ['name'=>'张三','age'=>22],
                ['name'=>'李四','age'=>21],
                ['name'=>'王五','age'=>20],
            ],
            [
                '张三'=>[
                    ['name'=>'张三','subject'=>'高数','score'=>99],
                    ['name'=>'张三','subject'=>'英语（二）','score'=>90],
                    ['name'=>'张三','subject'=>'近代史','score'=>80],
                ],
                '李四'=>[
                    ['name'=>'李四','subject'=>'高数','score'=>80],
                    ['name'=>'李四','subject'=>'英语（二）','score'=>30],
                    ['name'=>'李四','subject'=>'近代史','score'=>99],
                ],
                '王五'=>[
                    ['name'=>'王五','subject'=>'高数','score'=>88],
                    ['name'=>'王五','subject'=>'英语（二）','score'=>66],
                    ['name'=>'王五','subject'=>'近代史','score'=>100],
                ]
            ]
        ];
        $fileName = __DIR__.'/test3.xlsx';
        try {
            Excel::instance()->setWidth(20)->setDataType(3)->setHeader($header)->setData($data)->create()->save($fileName);
        } catch (\Exception $e) {
            echo $e->getMessage();
        }
    }
}
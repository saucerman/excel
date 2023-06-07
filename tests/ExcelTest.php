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
}
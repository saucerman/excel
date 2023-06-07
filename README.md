# excel

#### 介绍
对导出excel表格的封装处理

#### 安装教程
```sh
composer require byk/excel
```

#### 使用说明

1.  导出为文件
```php
use Byk\Excel\Excel;
$headers = [
            ['field'=>'name','title'=>'姓名'],
            ['field'=>'age','title'=>'年龄'],
        ];
$data = [
    ['name'=>'张三','age'=>22],
];
// 保存的文件地址
$fileName = __DIR__.'/test.xls';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setHeader($headers)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->setFileType('xls')//设置导出的文件格式
    ->create()//生成数据
    ->save($fileName);//保存为文件
    // $res===true 保存成功，否则返回的是错误信息
} catch (\Exception $e) {
}
```
2.  导出文件并合并某列
```php
use Byk\Excel\Excel;
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
    ]
];
// 保存的文件地址
$fileName = __DIR__.'/test.xls';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(2)//设置填充数据类型，1.普通二维数组，2.用户合并某列的三维数组
    ->setHeader($headers)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->setFileType('xls')//设置导出的文件格式
    ->create()//生成数据
    ->save($fileName);//保存为文件
    // $res===true 保存成功，否则返回的是错误信息
} catch (\Exception $e) {
}
```
3.  下载文件
```php
use Byk\Excel\Excel;
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
    ]
];
// 下载的文件名
$fileName = 'test.xls';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(2)//设置填充数据类型，1.普通二维数组，2.用户合并某列的三维数组
    ->setHeader($headers)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->setFileType('xls')//设置导出的文件格式
    ->create()//生成数据
    ->download($fileName);//下载
    // $res===true 保存成功，否则返回的是错误信息
} catch (\Exception $e) {
}
``` 
4. 自定义处理数据
```php
use Byk\Excel\Excel;
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
    ]
];
// 下载的文件名
$fileName = 'test.xls';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(2)//设置填充数据类型，1.普通二维数组，2.用户合并某列的三维数组
    ->setHeader($headers)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->setFileType('xls')//设置导出的文件格式
    ->create(function (\PhpOffice\PhpSpreadsheet\Spreadsheet $excel,$header, $data){
        //处理逻辑
    })//生成数据
    ->download($fileName);//下载
    // $res===true 保存成功，否则返回的是错误信息
} catch (\Exception $e) {
}
```
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
$fileName = '/path/test.xls';
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
$fileName = '/path/test.xls';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(2)//设置填充数据类型，1.普通二维数组，2.用于合并某列的三维数组 3.多个工作表数据
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
    Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(2)//设置填充数据类型，1.普通二维数组，2.用于合并某列的三维数组 3.多个工作表数据
    ->setHeader($headers)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->setFileType('xls')//设置导出的文件格式
    ->create()//生成数据
    ->download($fileName);//下载
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
// 保存的文件路径
$fileName = '/path/test.xls';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(2)//设置填充数据类型，1.普通二维数组，2.用于合并某列的三维数组 3.多个工作表数据
    ->setHeader($headers)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->setFileType('xls')//设置导出的文件格式
    ->create(function (\PhpOffice\PhpSpreadsheet\Spreadsheet $excel,$header, $data){
        //处理逻辑
    })//生成数据
    ->save($fileName);//保存为文件
    // $res===true 保存成功，否则返回的是错误信息
} catch (\Exception $e) {
}
```
5. 多个工作表
```php
use Byk\Excel\Excel;
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
$fileName = '/path/test.xlsx';
try {
    $res = Excel
    ::instance()//初始化
    ->setWidth(20)//设置单元格默认宽度
    ->setDataType(3)//设置填充数据类型，1.普通二维数组，2.用于合并某列的三维数组 3.多个工作表数据
    ->setHeader($header)//设置标题头数据
    ->setData($data)//设置要填充的数据
    ->create()//生成数据
    ->save($fileName);//保存为文件
    // $res===true 保存成功，否则返回的是错误信息
} catch (\Exception $e) {
}
```
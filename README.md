<h1 align="center">XlsxToCsv</h2=1>
<h4 align="center">
    快速把xlsx转换成csv文件
</h4>    

##

### 安装
本库依赖php扩展xlswriter
首先要先安装扩展xlswriter

```
https://xlswriter-docs.viest.me/zh-cn/an-zhuang
```
安装完成之后，安装本库
```
composer require --dev lovefc/xlsxtocsv
```

## 使用方法
创建使用文件run.php

```
<?php

require __DIR__ . '/vendor/autoload.php';

$config = ['path' => 'D:/xlsx', 'output' => 'D:/csv'];

$obj = new lovefc\XlsxToCsv($config);

$obj->run();

```
接着再去用命令行执行本文件


```
php run.php
```

## config配置
|   参数名称  |   说明  |
| --- | --- |
|  path  |    要读取的xlsx存在的目录或者xlsx文件  |
|  output  |   转化之后的输出目录,默认为xlsx文件所在的目录(可选,绝对路径) |
|  outputFile  |   转化之后的输出文件,会将所有文件所有分片归到一个文件(可选,绝对路径) |
|  reserves |    只保留的字段名,默认为空  |
|  showLog  |   输出日志，默认为true |
|  writeNum  |    每次写入数量，默认为5000 |
|  auto  |    是否自动写入csv,默认为true |
|  callBack |    回调函数，拥有两个参数，可在里面处理自己的数组逻辑  |
|  type |  类型转化数组 |
|  autoDeleSourceFile  |    转化完成之后,是否自动删除源文件,默认为false |

> type 类型转化，具体请参考这里：https://xlswriter-docs.viest.me/zh-cn/reader/set-type

## callBack的默认回调如下：
```
// 只提取手机号
$callback = function ($row, $sheetName) {
    $str = implode(',', $row);
	$pattern = '/1[0-9]\d{9}/';
	preg_match($pattern,$str,$arr);	
	$phone = $arr[0] ?? 0;
	if($phone == 0){
		return false;
	}	
    return $phone;
};

$config = ['path' => '/www/1.xlsx', 'outputFile' => '/www/1.csv', 'callBack'=>$callback ];

$obj = new lovefc\XlsxToCsv($config);

$obj->run();
 ```                    
> row是一个读取的数组，sheetName是当前的工作表名称

## License

[MIT](https://opensource.org/licenses/MIT)

Copyright (c) 2023-[lovefc](http://lovefc.cn)
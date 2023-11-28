<h1 align="center">XlsxToCsv</h2=1>
<h4 align="center">
    快速把xlsx转换成csv文件
</h4>    

##

### 安装
```
composer require --dev lovefc/xlsxtocsv
```

## 使用方法

```

$config = ['path' => 'D:/xlsx', 'output' => 'D:/csv'];

$obj = new lovefc\XlsxToCsv($config);

$obj->run();

```

## config配置
|   参数名称  |   说明  |
| --- | --- |
|  path  |    要读取的xlsx存在的目录或者xlsx文件  |
|  output  |   转化之后的输出目录,默认为xlsx文件所在的目录 |
|  reserves |    只保留的字段名,默认为空  |
|  showLog  |   输出日志，默认为true |
|  writeNum  |    每次写入数量，默认为5000 |
|  auto  |    是否自动写入csv,默认为true |
|  callBack |    回调函数，拥有两个参数，可在里面处理自己的数组逻辑  |
|  type |  类型转化数组 |

> type 类型转化，具体请参考这里：https://xlswriter-docs.viest.me/zh-cn/reader/set-type

## callBack的默认回调如下：
```
$callback = function ($row, $sheetName) {
     $text = implode(',', $row);
     return $text;
 };
 ```                    
> row是一个读取的数组，sheetName是当前的工作表名称

## License

[MIT](https://opensource.org/licenses/MIT)

Copyright (c) 2023-[lovefc](http://lovefc.cn)
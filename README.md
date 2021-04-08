## ChenExcel

[![Build Status](https://github.com/chenmobuys/excel/workflows/master/badge.svg)](https://github.com/chenmobuys/excel/actions)
[![Latest Stable Version](https://img.shields.io/packagist/v/chen/excel.svg)](https://packagist.org/packages/chen/excel) 
[![Total Downloads](https://img.shields.io/packagist/dt/chen/excel)](https://packagist.org/packages/chen/excel) 
[![License](https://img.shields.io/packagist/l/chen/excel)](https://packagist.org/packages/chen/excel) 
[![Platform Support](https://img.shields.io/packagist/php-v/chen/excel)](https://github.com/chenmobuys/excel)

## 描述
ChenExcel 主要目的是有效的读取表格中的数据，可以处理大型文件，它可能不是效率最高的，但至少不会耗尽内存

目前支持 CSV、ODS、XLSX、XLS，XLS使用的是 [https://code.google.com/archive/p/php-excel-reader/](https://code.google.com/archive/p/php-excel-reader/)，较大的表格依然存在问题，因为它是一次性读取所有数据，然后将所有内容留在内存中

## 环境

1. php 版本 ^5.6||^7.0||~8.0,<8.1
3. Zip扩展
2. [Composer](https://getcomposer.org/)

## 安装
```bash
composer require chen/excel -vvv
```

## 用法
基本使用

```php
<?php

use Chen\Excel\SpreadsheetReader;

// 根据文件后缀名识别文件类型
$filename = 'sample.csv';
$reader = SpreadsheetReader::load($filename);

// 获取表格数据，默认为第一个表格
$rows = [];
foreach($reader as $row) {
    $rows[] = $row;
}

// 获取表格所有sheetName
$sheetNames = $reader->getSheetNames();

// 获取指定表格数据
$reader->setSheetIndex(1);
...

```

注册自定义Reader
```php
<?php

use Chen\Excel\Readers\BaseReader;

class SampleReader extends BaseReader {
    
    /**
     * Load handle from filename
     *
     * @return void
     */
    public function load($filename, $options): void 
    {
        // TODO
    }
    ...
}

----------------------------------------------

<?php

use SampleReader;
use Chen\Excel\SpreadsheetReader;

// 自定义文件读取器
$filename = 'sample.xls';

SpreadsheetReader::registerReader('Xls', SampleReader::class);

$reader = SpreadsheetReader::load($filename);

```

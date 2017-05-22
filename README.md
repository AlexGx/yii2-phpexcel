PHPExcel extension for Yii2
=============

[![Latest Stable Version](https://poser.pugx.org/alexgx/yii2-phpexcel/v/stable.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel) [![Total Downloads](https://poser.pugx.org/alexgx/yii2-phpexcel/downloads.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel) [![Latest Unstable Version](https://poser.pugx.org/alexgx/yii2-phpexcel/v/unstable.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel) [![License](https://poser.pugx.org/alexgx/yii2-phpexcel/license.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel)

%short_description%

Installation
------------
The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require "alexgx/yii2-phpexcel" "*"
```
or add

```json
"alexgx/yii2-phpexcel" : "*"
```

to the require section of your application's `composer.json` file.

Features
---------------
TDB


Usage
-----
```php

use alexgx\phpexcel\PhpExcel;

$phpExcel = new PhpExcel();
$objPHPExcel = $phpExcel->create();

$objPHPExcel->getProperties()->setCreator("Traiding")
    ->setLastModifiedBy("Uldis Nelsons")
    ->setTitle("Packing list")
    ->setSubject("Packing list");

$activeSheet = $objPHPExcel->setActiveSheetIndex(0);
$activeSheet->setTitle('Packing List')
    ->setCellValue('A1', 'PACKING LIST')
    ->setCellValue('A3', 'VESSEL:')
    ->setCellValue('A4', 'B/L date:')
    ->setCellValue('A5', 'B/L No.:');



$activeSheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
$activeSheet->getStyle('A3:c9')->getFont()->setBold(true)->setSize(11);

$writer = new ExcelDataWriter();
$writer->setStartRow(11);
$writer->sheet = $activeSheet;
$writer->data = $packingList;

$headerStyles = [
    'font' => [
        'bold' => true
    ]
];
$writer->columns = [
    [
        'attribute' => 'departure_date',
        'header' => 'Departure Date',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'car_number',
        'header' => 'Car Number',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'delivery_note',
        'header' => 'Delivery Note',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'weight',
        'header' => 'Weight',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'gtd',
        'header' => 'Gtd',
        'headerStyles' => $headerStyles,
    ]
];

$writer->write();
$phpExcel->responseFile($objPHPExcel, 'packing.xls');


```
TBD

Further Information
-------------------
Please, check the [PHPOffice/PHPExcel github repo](https://github.com/PHPOffice/PHPExcel) documentation for further information about its configuration options.

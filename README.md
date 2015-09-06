Yii2 PHP Excel
==============

[![Latest Stable Version](https://poser.pugx.org/moonlandsoft/yii2-phpexcel/v/stable)](https://packagist.org/packages/moonlandsoft/yii2-phpexcel) 
[![Total Downloads](https://poser.pugx.org/moonlandsoft/yii2-phpexcel/downloads)](https://packagist.org/packages/moonlandsoft/yii2-phpexcel) 
[![Latest Unstable Version](https://poser.pugx.org/moonlandsoft/yii2-phpexcel/v/unstable)](https://packagist.org/packages/moonlandsoft/yii2-phpexcel) 
[![License](https://poser.pugx.org/moonlandsoft/yii2-phpexcel/license)](https://packagist.org/packages/moonlandsoft/yii2-phpexcel)

Exporting PHP to Excel or Importing Excel to PHP.
Excel Widget for generate Excel File or for load Excel File.


Property
--------

string `$mode` is an export mode or import mode. valid value are 'export' and 'import'  

boolean `$isMultipleSheet` for set the export excel with multiple sheet.  

array `$properties` for set property on the excel object.  

array `$models` Model object or DataProvider object with much data.  

array `$columns` to get the attributes from the model, this valid value only the exist attribute on the model. If this is not set, then all attribute of the model will be set as columns.  

array `$headers` to set the header column on first line. Set this if want to custom header. If not set, the header will get attributes label of model attributes.  

string|array `$fileName` is a name for file name to export or import. Multiple file name only use for import mode, not work if you use the export mode.  

string `$savePath` is a directory to save the file or you can blank this to set the file as attachment.  

string `$format` for excel to export. Valid value are 'Excel5', 'Excel2007', 'Excel2003XML', '00Calc', 'Gnumeric'.  

boolean `$setFirstTitle` to set the title column on the first line. The columns will have a header on the first line.  

boolean `$asAttachment` to set the file excel to download mode.  

boolean `$setFirstRecordAsKeys` to set the first record on excel file to a keys of array per line. If you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.  

boolean `$setIndexSheetByName` to set the sheet index by sheet name or array result if the sheet not only one  

string `$getOnlySheet` is a sheet name to getting the data. This is only get the sheet with same name.

array|Formatter `$formatter` the formatter used to format model attribute values into displayable texts. This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]] instance. If this property is not set, the "formatter" application component will be used.

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist moonlandsoft/yii2-phpexcel "*"
```

or add

```
"moonlandsoft/yii2-phpexcel": "*"
```

to the require section of your `composer.json` file.


Usage 
-----

### Exporting Data

Exporting data into an excel file.


```php
<?php

// export data only one worksheet.

\moonland\phpexcel\Excel::widget([
	'models' => $allModels,
	'mode' => 'export', //default value as 'export'
	'columns' => ['column1','column2','column3'], //without header working, because the header will be get label from attribute label. 
	'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'], 
]);

\moonland\phpexcel\Excel::export([
	'models' => $allModels, 
	'columns' => ['column1','column2','column3'], //without header working, because the header will be get label from attribute label. 
	'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
]);

// export data with multiple worksheet.

\moonland\phpexcel\Excel::widget([
	'isMultipleSheet' => true, 
	'models' => [
		'sheet1' => $allModels1, 
		'sheet2' => $allModels2, 
		'sheet3' => $allModels3
	], 
	'mode' => 'export', //default value as 'export' 
	'columns' => [
		'sheet1' => ['column1','column2','column3'], 
		'sheet2' => ['column1','column2','column3'], 
		'sheet3' => ['column1','column2','column3']
	],
	//without header working, because the header will be get label from attribute label. 
	'header' => [
		'sheet1' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'], 
		'sheet2' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'], 
		'sheet3' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3']
	],
]);

\moonland\phpexcel\Excel::export([
	'isMultipleSheet' => true, 
	'models' => [
		'sheet1' => $allModels1, 
		'sheet2' => $allModels2, 
		'sheet3' => $allModels3
	], 'columns' => [
		'sheet1' => ['column1','column2','column3'], 
		'sheet2' => ['column1','column2','column3'], 
		'sheet3' => ['column1','column2','column3']
	], 
	//without header working, because the header will be get label from attribute label. 
	'header' => [
		'sheet1' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
		'sheet2' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
		'sheet3' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3']
	],
]);

```

New Feature for exporting data, you can use this if you familiar yii gridview. 
That is same with gridview data column.
Columns in array mode valid params are 'attribute', 'header', 'format', 'value', and footer (TODO).
Columns in string mode valid layout are 'attribute:format:header:footer(TODO)'.
  
```php
<?php
  
\moonland\phpexcel\Excel::export([
   	'models' => Post::find()->all(),
      	'columns' => [
      		'author.name:text:Author Name',
      		[
      				'attribute' => 'content',
      				'header' => 'Content Post',
      				'format' => 'text',
      				'value' => function($model) {
      					return ExampleClass::removeText('example', $model->content);
      				},
      		],
      		'like_it:text:Reader like this content',
      		'created_at:datetime',
      		[
      				'attribute' => 'updated_at',
      				'format' => 'date',
      		],
      	],
      	'headers' => [
     		'created_at' => 'Date Created Content',
		],
]);
	  
```


### Importing Data

Import file excel and return into an array.


```php
<?php

$data = \moonland\phpexcel\Excel::import($fileName, $config); // $config is an optional

$data = \moonland\phpexcel\Excel::widget([
		'mode' => 'import', 
		'fileName' => $fileName, 
		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel. 
		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric. 
		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
	]);

$data = \moonland\phpexcel\Excel::import($fileName, [
		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel. 
		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric. 
		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
	]);

// import data with multiple file.

$data = \moonland\phpexcel\Excel::widget([
	'mode' => 'import', 
	'fileName' => [
		'file1' => $fileName1, 
		'file2' => $fileName2, 
		'file3' => $fileName3,
	], 
		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel. 
		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric. 
		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
	]);

$data = \moonland\phpexcel\Excel::import([
	'file1' => $fileName1, 
	'file2' => $fileName2, 
	'file3' => $fileName3,
	], [
		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel. 
		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric. 
		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
	]);

```

Result example from the code on the top :

```

// only one sheet or specified sheet.
Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2));

// data with multiple worksheet
Array([Sheet1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)), [Sheet2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)));

// data with multiple file and specified sheet or only one worksheet
Array([file1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)), [file2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)));

// data with multiple file and multiple worksheet
Array([file1] => Array([Sheet1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)), [Sheet2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2))), [file2] => Array([Sheet1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)), [Sheet2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2), [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2))));

```

TODO
----
- Adding footer params for columns in exporting data.

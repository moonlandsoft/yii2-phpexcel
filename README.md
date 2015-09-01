Yii2 PHP Excel
==============
Exporting PHP to Excel

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

Once the extension is installed, simply use it in your code by  :

```php
<?php
\moonland\phpexcel\Excel::widget([
	'models' => $allModels,
	'columns' => ['column1','column2','column3'],
	'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
]); ?>
```
Exporting data into an excel file.

```php
<?php
\moonland\phpexcel\Excel::export([
 		'models' => $allModels,
 		'columns' => ['column1','column2','column3'],
 		//without header working, because the header will be get label from attribute label.
 		'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
]);
 
```

Import file excel and return into an array.
```php
<?php

$data = \moonland\phpexcel\Excel::import($fileName, $config); // $config is an optional

$data = \moonland\phpexcel\Excel::import($fileName, ['setFirstRecordAsKeys' => true]);

```
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
<?= \moonland\phpexcel\Excel::widget([
	'models' => $allModels,
	'columns' => ['column1','column2','column3'],
	'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
]); ?>
```
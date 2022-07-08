<?php

namespace moonland\phpexcel;

use yii\helpers\ArrayHelper;
use yii\base\InvalidConfigException;
use yii\base\InvalidParamException;
use yii\i18n\Formatter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Excel Widget for generate Excel File or for load Excel File.
 *
 * Usage
 * -----
 *
 * Exporting data into an excel file.
 *
 * ~~~
 *
 * // export data only one worksheet.
 *
 * \moonland\phpexcel\Excel::widget([
 * 		'models' => $allModels,
 * 		'mode' => 'export', //default value as 'export'
 * 		'columns' => ['column1','column2','column3'],
 * 		//without header working, because the header will be get label from attribute label.
 * 		'headers' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * ]);
 *
 * \moonland\phpexcel\Excel::export([
 * 		'models' => $allModels,
 * 		'columns' => ['column1','column2','column3'],
 * 		//without header working, because the header will be get label from attribute label.
 * 		'headers' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * ]);
 *
 * // export data with multiple worksheet.
 *
 * \moonland\phpexcel\Excel::widget([
 * 		'isMultipleSheet' => true,
 * 		'models' => [
 * 			'sheet1' => $allModels1,
 * 			'sheet2' => $allModels2,
 * 			'sheet3' => $allModels3
 * 		],
 * 		'mode' => 'export', //default value as 'export'
 * 		'columns' => [
 * 			'sheet1' => ['column1','column2','column3'],
 * 			'sheet2' => ['column1','column2','column3'],
 * 			'sheet3' => ['column1','column2','column3']
 * 		],
 * 		//without header working, because the header will be get label from attribute label.
 * 		'headers' => [
 * 			'sheet1' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * 			'sheet2' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * 			'sheet3' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3']
 * 		],
 * ]);
 *
 * \moonland\phpexcel\Excel::export([
 * 		'isMultipleSheet' => true,
 * 		'models' => [
 * 			'sheet1' => $allModels1,
 * 			'sheet2' => $allModels2,
 * 			'sheet3' => $allModels3
 * 		],
 * 		'columns' => [
 * 			'sheet1' => ['column1','column2','column3'],
 * 			'sheet2' => ['column1','column2','column3'],
 * 			'sheet3' => ['column1','column2','column3']
 * 		],
 * 		//without header working, because the header will be get label from attribute label.
 * 		'headers' => [
 * 			'sheet1' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * 			'sheet2' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * 			'sheet3' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3']
 * 		],
 * ]);
 *
 * ~~~
 *
 * New Feature for exporting data, you can use this if you familiar yii gridview.
 * That is same with gridview data column.
 * Columns in array mode valid params are 'attribute', 'header', 'format', 'value', and footer (TODO).
 * Columns in string mode valid layout are 'attribute:format:header:footer(TODO)'.
 *
 * ~~~
 *
 * \moonland\phpexcel\Excel::export([
 *  	'models' => Post::find()->all(),
 *     	'columns' => [
 *     		'author.name:text:Author Name',
 *     		[
 *     				'attribute' => 'content',
 *     				'header' => 'Content Post',
 *     				'format' => 'text',
 *     				'value' => function($model) {
 *     					return ExampleClass::removeText('example', $model->content);
 *     				},
 *     		],
 *     		'like_it:text:Reader like this content',
 *     		'created_at:datetime',
 *     		[
 *     				'attribute' => 'updated_at',
 *     				'format' => 'date',
 *     		],
 *     	],
 *     	'headers' => [
 *     		'created_at' => 'Date Created Content',
 * 		],
 * ]);
 *
 * ~~~
 *
 *
 * Import file excel and return into an array.
 *
 * ~~~
 *
 * $data = \moonland\phpexcel\Excel::import($fileName, $config); // $config is an optional
 *
 * $data = \moonland\phpexcel\Excel::widget([
 * 		'mode' => 'import',
 * 		'fileName' => $fileName,
 * 		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
 * 		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric.
 * 		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
 * ]);
 *
 * $data = \moonland\phpexcel\Excel::import($fileName, [
 * 		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
 * 		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric.
 * 		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
 *	]);
 *
 * // import data with multiple file.
 *
 * $data = \moonland\phpexcel\Excel::widget([
 * 		'mode' => 'import',
 * 		'fileName' => [
 * 			'file1' => $fileName1,
 * 			'file2' => $fileName2,
 * 			'file3' => $fileName3,
 * 		],
 * 		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
 * 		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric.
 * 		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
 * ]);
 *
 * $data = \moonland\phpexcel\Excel::import([
 * 			'file1' => $fileName1,
 * 			'file2' => $fileName2,
 * 			'file3' => $fileName3,
 * 		], [
 * 		'setFirstRecordAsKeys' => true, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
 * 		'setIndexSheetByName' => true, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric.
 * 		'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
 *	]);
 *
 * ~~~
 *
 * Result example from the code on the top :
 *
 * ~~~
 *
 * // only one sheet or specified sheet.
 *
 * Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2));
 *
 * // data with multiple worksheet
 *
 * Array([Sheet1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)),
 * [Sheet2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)));
 *
 * // data with multiple file and specified sheet or only one worksheet
 *
 * Array([file1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)),
 * [file2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)));
 *
 * // data with multiple file and multiple worksheet
 *
 * Array([file1] => Array([Sheet1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)),
 * [Sheet2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2))),
 * [file2] => Array([Sheet1] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2)),
 * [Sheet2] => Array([0] => Array([name] => Anam, [email] => moh.khoirul.anaam@gmail.com, [framework interest] => Yii2),
 * [1] => Array([name] => Example, [email] => example@moonlandsoft.com, [framework interest] => Yii2))));
 *
 * ~~~
 *
 * @property string $mode is an export mode or import mode. valid value are 'export' and 'import'
 * @property boolean $isMultipleSheet for set the export excel with multiple sheet.
 * @property array $properties for set property on the excel object.
 * @property array $models Model object or DataProvider object with much data.
 * @property array $columns to get the attributes from the model, this valid value only the exist attribute on the model.
 * If this is not set, then all attribute of the model will be set as columns.
 * @property array $headers to set the header column on first line. Set this if want to custom header.
 * If not set, the header will get attributes label of model attributes.
 * @property string|array $fileName is a name for file name to export or import. Multiple file name only use for import mode, not work if you use the export mode.
 * @property string $savePath is a directory to save the file or you can blank this to set the file as attachment.
 * @property string $format for excel to export. Valid value are 'Xls','Xlsx','Xml','Ods','Slk','Gnumeric','Csv', and 'Html'.
 * @property boolean $setFirstTitle to set the title column on the first line. The columns will have a header on the first line.
 * @property boolean $asAttachment to set the file excel to download mode.
 * @property boolean $setFirstRecordAsKeys to set the first record on excel file to a keys of array per line.
 * If you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
 * @property boolean $setIndexSheetByName to set the sheet index by sheet name or array result if the sheet not only one
 * @property string $getOnlySheet is a sheet name to getting the data. This is only get the sheet with same name.
 * @property array|Formatter $formatter the formatter used to format model attribute values into displayable texts.
 * This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]]
 * instance. If this property is not set, the "formatter" application component will be used.
 *
 * @author Moh Khoirul Anam <moh.khoirul.anaam@gmail.com>
 * @copyright 2014
 * @since 1
 */
class Excel extends \yii\base\Widget
{
    // Border style
    const BORDER_NONE = 'none';
    const BORDER_DASHDOT = 'dashDot';
    const BORDER_DASHDOTDOT = 'dashDotDot';
    const BORDER_DASHED = 'dashed';
    const BORDER_DOTTED = 'dotted';
    const BORDER_DOUBLE = 'double';
    const BORDER_HAIR = 'hair';
    const BORDER_MEDIUM = 'medium';
    const BORDER_MEDIUMDASHDOT = 'mediumDashDot';
    const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    const BORDER_MEDIUMDASHED = 'mediumDashed';
    const BORDER_SLANTDASHDOT = 'slantDashDot';
    const BORDER_THICK = 'thick';
    const BORDER_THIN = 'thin';
    
    // Colors
    const COLOR_BLACK = 'FF000000';
    const COLOR_WHITE = 'FFFFFFFF';
    const COLOR_RED = 'FFFF0000';
    const COLOR_DARKRED = 'FF800000';
    const COLOR_BLUE = 'FF0000FF';
    const COLOR_DARKBLUE = 'FF000080';
    const COLOR_GREEN = 'FF00FF00';
    const COLOR_DARKGREEN = 'FF008000';
    const COLOR_YELLOW = 'FFFFFF00';
    const COLOR_DARKYELLOW = 'FF808000';
    
    // Horizontal alignment styles
    const HORIZONTAL_GENERAL = 'general';
    const HORIZONTAL_LEFT = 'left';
    const HORIZONTAL_RIGHT = 'right';
    const HORIZONTAL_CENTER = 'center';
    const HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous';
    const HORIZONTAL_JUSTIFY = 'justify';
    const HORIZONTAL_FILL = 'fill';
    const HORIZONTAL_DISTRIBUTED = 'distributed'; // Excel2007 only
    
    // Vertical alignment styles
    const VERTICAL_BOTTOM = 'bottom';
    const VERTICAL_TOP = 'top';
    const VERTICAL_CENTER = 'center';
    const VERTICAL_JUSTIFY = 'justify';
    const VERTICAL_DISTRIBUTED = 'distributed'; // Excel2007 only
    
    // Read order
    const READORDER_CONTEXT = 0;
    const READORDER_LTR = 1;
    const READORDER_RTL = 2;
    
    // Fill types
    const FILL_NONE = 'none';
    const FILL_SOLID = 'solid';
    const FILL_GRADIENT_LINEAR = 'linear';
    const FILL_GRADIENT_PATH = 'path';
    const FILL_PATTERN_DARKDOWN = 'darkDown';
    const FILL_PATTERN_DARKGRAY = 'darkGray';
    const FILL_PATTERN_DARKGRID = 'darkGrid';
    const FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal';
    const FILL_PATTERN_DARKTRELLIS = 'darkTrellis';
    const FILL_PATTERN_DARKUP = 'darkUp';
    const FILL_PATTERN_DARKVERTICAL = 'darkVertical';
    const FILL_PATTERN_GRAY0625 = 'gray0625';
    const FILL_PATTERN_GRAY125 = 'gray125';
    const FILL_PATTERN_LIGHTDOWN = 'lightDown';
    const FILL_PATTERN_LIGHTGRAY = 'lightGray';
    const FILL_PATTERN_LIGHTGRID = 'lightGrid';
    const FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal';
    const FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis';
    const FILL_PATTERN_LIGHTUP = 'lightUp';
    const FILL_PATTERN_LIGHTVERTICAL = 'lightVertical';
    const FILL_PATTERN_MEDIUMGRAY = 'mediumGray';
    
    // Pre-defined formats
    const FORMAT_GENERAL = 'General';
    
    const FORMAT_TEXT = '@';
    
    const FORMAT_NUMBER = '0';
    const FORMAT_NUMBER_00 = '0.00';
    const FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00';
    const FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-';
    
    const FORMAT_PERCENTAGE = '0%';
    const FORMAT_PERCENTAGE_00 = '0.00%';
    
    const FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd';
    const FORMAT_DATE_YYYYMMDD = 'yy-mm-dd';
    const FORMAT_DATE_DDMMYYYY = 'dd/mm/yy';
    const FORMAT_DATE_DMYSLASH = 'd/m/yy';
    const FORMAT_DATE_DMYMINUS = 'd-m-yy';
    const FORMAT_DATE_DMMINUS = 'd-m';
    const FORMAT_DATE_MYMINUS = 'm-yy';
    const FORMAT_DATE_XLSX14 = 'mm-dd-yy';
    const FORMAT_DATE_XLSX15 = 'd-mmm-yy';
    const FORMAT_DATE_XLSX16 = 'd-mmm';
    const FORMAT_DATE_XLSX17 = 'mmm-yy';
    const FORMAT_DATE_XLSX22 = 'm/d/yy h:mm';
    const FORMAT_DATE_DATETIME = 'd/m/yy h:mm';
    const FORMAT_DATE_TIME1 = 'h:mm AM/PM';
    const FORMAT_DATE_TIME2 = 'h:mm:ss AM/PM';
    const FORMAT_DATE_TIME3 = 'h:mm';
    const FORMAT_DATE_TIME4 = 'h:mm:ss';
    const FORMAT_DATE_TIME5 = 'mm:ss';
    const FORMAT_DATE_TIME6 = 'h:mm:ss';
    const FORMAT_DATE_TIME7 = 'i:s.S';
    const FORMAT_DATE_TIME8 = 'h:mm:ss;@';
    const FORMAT_DATE_YYYYMMDDSLASH = 'yy/mm/dd;@';
    
    const FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-';
    const FORMAT_CURRENCY_USD = '$#,##0_-';
    const FORMAT_CURRENCY_EUR_SIMPLE = '#,##0.00_-"€"';
    const FORMAT_CURRENCY_EUR = '#,##0_-"€"';
    
	/**
	 * @var string mode is an export mode or import mode. valid value are 'export' and 'import'.
	 */
	public $mode = 'export';
	/**
	 * @var boolean for set the export excel with multiple sheet.
	 */
	public $isMultipleSheet = false;
	/**
	 * @var array properties for set property on the excel object.
	 */
	public $properties;
	/**
	 * @var Model object or DataProvider object with much data.
	 */
	public $models;
	/**
	 * @var array columns to get the attributes from the model, this valid value only the exist attribute on the model.
	 * If this is not set, then all attribute of the model will be set as columns.
	 */
	public $columns = [];
	/**
	 * @var array header to set the header column on first line. Set this if want to custom header.
	 * If not set, the header will get attributes label of model attributes.
	 */
	public $headers = [];
	/**
	 * @var string|array name for file name to export or save.
	 */
	public $fileName;
	/**
	 * @var string save path is a directory to save the file or you can blank this to set the file as attachment.
	 */
	public $savePath;
	/**
	 * @var string format for excel to export. Valid value are 'Xls','Xlsx','Xml','Ods','Slk','Gnumeric','Csv', and 'Html'.
	 */
	public $format;
	/**
	 * @var boolean to set the title column on the first line.
	 */
	public $setFirstTitle = true;
	/**
	 * @var boolean to set the file excel to download mode.
	 */
	public $asAttachment = false;
	/**
	 * @var boolean to set the first record on excel file to a keys of array per line.
	 * If you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
	 */
	public $setFirstRecordAsKeys = true;
	/**
	 * @var boolean to set the sheet index by sheet name or array result if the sheet not only one.
	 */
	public $setIndexSheetByName = false;
	/**
	 * @var string sheetname to getting. This is only get the sheet with same name.
	 */
	public $getOnlySheet;
	/**
	 * @var boolean to set the import data will return as array.
	 */
	public $asArray;
	/**
	 * @var array to unread record by index number.
	 */
	public $leaveRecordByIndex = [];
	/**
	 * @var array to read record by index, other will leave.
	 */
	public $getOnlyRecordByIndex = [];
	/**
	 * @var array|Formatter the formatter used to format model attribute values into displayable texts.
	 * This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]]
	 * instance. If this property is not set, the "formatter" application component will be used.
	 */
	public $formatter;
	/**
	 * @var boolean define the column autosize
	 */
	public $autoSize = false;
	/**
	 * @var boolean if true, this writer pre-calculates all formulas in the spreadsheet. This can be slow on large spreadsheets, and maybe even unwanted.
	 */
	public $preCalculationFormula = false;
	/**
	 * @var boolean Because of a bug in the Office2003 compatibility pack, there can be some small issues when opening Xlsx spreadsheets (mostly related to formula calculation)
	 */
	public $compatibilityOffice2003 = false;
	/**
	 * @var custom CSV delimiter for import. Works only with CSV files
	 */
	public $CSVDelimiter = ";";
	/**
	 * @var custom CSV encoding for import. Works only with CSV files
	 */
	public $CSVEncoding = "UTF-8";
  
	/**
	 * (non-PHPdoc)
	 * @see \yii\base\Object::init()
	 */
	public function init()
	{
		parent::init();
		if ($this->formatter == null) {
			$this->formatter = \Yii::$app->getFormatter();
		} elseif (is_array($this->formatter)) {
			$this->formatter = \Yii::createObject($this->formatter);
		}
		if (!$this->formatter instanceof Formatter) {
			throw new InvalidConfigException('The "formatter" property must be either a Format object or a configuration array.');
		}
	}

	/**
	 * Setting data from models
	 */
	public function executeColumns($models, &$activeSheet = null,  $columns = [], $headers = [])
	{
		if ($activeSheet == null) {
			$activeSheet = $this->activeSheet;
		}
		$hasHeader = false;
		$row = 1;
		$char = 26;
		foreach ($models as $model) {
			if (empty($columns)) {
				$columns = $model->attributes();
			}
			if ($this->setFirstTitle && !$hasHeader) {
				$isPlus = false;
				$colplus = 0;
				$colnum = 1;
				foreach ($columns as $key=>$column) {
					$col = '';
					if ($colnum > $char) {
						$colplus += 1;
						$colnum = 1;
						$isPlus = true;
					}
					if ($isPlus) {
						$col .= chr(64+$colplus);
					}
					$col .= chr(64+$colnum);
					$header = '';
					if (is_array($column)) {
						if (isset($column['header'])) {
							$header = $column['header'];
						} elseif (isset($column['attribute']) && isset($headers[$column['attribute']])) {
							$header = $headers[$column['attribute']];
						} elseif (isset($column['attribute'])) {
							$header = $model->getAttributeLabel($column['attribute']);
						} elseif (isset($column['cellFormat']) && is_array($column['cellFormat'])) {
						    $activeSheet->getStyle($col.$row)->applyFromArray($column['cellFormat']);
						}
					} else {
					    if(isset($headers[$column])) {
					        $header = $headers[$column];
					    } else {
					        $header = $model->getAttributeLabel($column);
					    }
					}
					if (isset($column['width'])) {
					    $activeSheet->getColumnDimension(strtoupper($col))->setWidth($column['width']);
					}
					$activeSheet->setCellValue($col.$row,$header);
					$colnum++;
				}
				$hasHeader=true;
				$row++;
			}
			$isPlus = false;
			$colplus = 0;
			$colnum = 1;
			foreach ($columns as $key=>$column) {
				$col = '';
				if ($colnum > $char) {
					$colplus++;
					$colnum = 1;
					$isPlus = true;
				}
				if ($isPlus) {
					$col .= chr(64+$colplus);
				}
				$col .= chr(64+$colnum);
				if (is_array($column)) {
				    $column_value = $this->executeGetColumnData($model, $column);
				    if (isset($column['cellFormat']) && is_array($column['cellFormat'])) {
				        $activeSheet->getStyle($col.$row)->applyFromArray($column['cellFormat']);
				    }
				} else {
					$column_value = $this->executeGetColumnData($model, ['attribute' => $column]);
				}
				$activeSheet->setCellValue($col.$row,$column_value);
				$colnum++;
			}
			$row++;
			
			if($this->autoSize){
				foreach (range(0, $colnum) as $col) {
					$activeSheet->getColumnDimensionByColumn($col)->setAutoSize(true);
				}
			}
		}
	}

	/**
	 * Setting label or keys on every record if setFirstRecordAsKeys is true.
	 * @param array $sheetData
	 * @return multitype:multitype:array
	 */
	public function executeArrayLabel($sheetData)
	{
		$keys = ArrayHelper::remove($sheetData, '1');

		$new_data = [];

		foreach ($sheetData as $values)
		{
			$new_data[] = array_combine($keys, $values);
		}

		return $new_data;
	}

	/**
	 * Leave record with same index number.
	 * @param array $sheetData
	 * @param array $index
	 * @return array
	 */
	public function executeLeaveRecords($sheetData = [], $index = [])
	{
		foreach ($sheetData as $key => $data)
		{
			if (in_array($key, $index))
			{
				unset($sheetData[$key]);
			}
		}
		return $sheetData;
	}

	/**
	 * Read record with same index number.
	 * @param array $sheetData
	 * @param array $index
	 * @return array
	 */
	public function executeGetOnlyRecords($sheetData = [], $index = [])
	{
		foreach ($sheetData as $key => $data)
		{
			if (!in_array($key, $index))
			{
				unset($sheetData[$key]);
			}
		}
		return $sheetData;
	}

	/**
	 * Getting column value.
	 * @param Model $model
	 * @param array $params
	 * @return Ambigous <NULL, string, mixed>
	 */
	public function executeGetColumnData($model, $params = [])
	{
		$value = null;
		if (isset($params['value']) && $params['value'] !== null) {
			if (is_string($params['value'])) {
				$value = ArrayHelper::getValue($model, $params['value']);
			} else {
				$value = call_user_func($params['value'], $model, $this);
			}
		} elseif (isset($params['attribute']) && $params['attribute'] !== null) {
			$value = ArrayHelper::getValue($model, $params['attribute']);
		}

		if (isset($params['format']) && $params['format'] != null)
			$value = $this->formatter()->format($value, $params['format']);

		return $value;
	}

	/**
	 * Populating columns for checking the column is string or array. if is string this will be checking have a formatter or header.
	 * @param array $columns
	 * @throws InvalidParamException
	 * @return multitype:multitype:array
	 */
	public function populateColumns($columns = [])
	{
		$_columns = [];
		foreach ($columns as $key => $value)
		{
			if (is_string($value))
			{
				$value_log = explode(':', $value);
				$_columns[$key] = ['attribute' => $value_log[0]];

				if (isset($value_log[1]) && $value_log[1] !== null) {
					$_columns[$key]['format'] = $value_log[1];
				}

				if (isset($value_log[2]) && $value_log[2] !== null) {
					$_columns[$key]['header'] = $value_log[2];
				}
			} elseif (is_array($value)) {
				if (!isset($value['attribute']) && !isset($value['value'])) {
					throw new \InvalidArgumentException('Attribute or Value must be defined.');
				}
				$_columns[$key] = $value;
			}
		}

		return $_columns;
	}

	/**
	 * Formatter for i18n.
	 * @return Formatter
	 */
	public function formatter()
	{
		if (!isset($this->formatter))
			$this->formatter = \Yii::$app->getFormatter();

		return $this->formatter;
	}

	/**
	 * Setting header to download generated file xls
	 */
	public function setHeaders()
	{
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="' . $this->getFileName() .'"');
		header('Cache-Control: max-age=0');
	}

	/**
	 * Getting the file name of exporting xls file
	 * @return string
	 */
	public function getFileName()
	{
		$fileName = 'exports.xlsx';
		if (isset($this->fileName)) {
		    $fileName = $this->fileName;
			if (strpos($fileName, '.xlsx') === false)
				$fileName .= '.xlsx';
		}
		return $fileName;
	}

	/**
	 * Setting properties for excel file
	 * @param PHPExcel $objectExcel
	 * @param array $properties
	 */
	public function properties(&$objectExcel, $properties = [])
	{
		foreach ($properties as $key => $value)
		{
			$keyname = "set" . ucfirst($key);
			$objectExcel->getProperties()->{$keyname}($value);
		}
	}

	/**
	 * saving the xls file to download or to path
	 */
	public function writeFile($sheet)
	{
		if (!isset($this->format))
			$this->format = 'Xlsx';
		$objectwriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($sheet, $this->format);
		$path = 'php://output';
		if (isset($this->savePath) && $this->savePath != null) {
			$path = $this->savePath . '/' . $this->getFileName();
		}
		$objectwriter->setOffice2003Compatibility($this->compatibilityOffice2003);
		$objectwriter->setPreCalculateFormulas($this->preCalculationFormula);
		$objectwriter->save($path);
		if ($path == 'php://output')
    		  exit();
		
    		 return true;
	}

	/**
	 * reading the xls file
	 */
	public function readFile($fileName)
	{
		if (!isset($this->format))
			$this->format = \PhpOffice\PhpSpreadsheet\IOFactory::identify($fileName);
		$objectreader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($this->format);
		if ($this->format == "Csv") {
			$objectreader->setDelimiter($this->CSVDelimiter);
			$objectreader->setInputEncoding($this->CSVEncoding);
		}
		$objectPhpExcel = $objectreader->load($fileName);

		$sheetCount = $objectPhpExcel->getSheetCount();

		$sheetDatas = [];

		if ($sheetCount > 1) {
			foreach ($objectPhpExcel->getSheetNames() as $sheetIndex => $sheetName) {
				if (isset($this->getOnlySheet) && $this->getOnlySheet != null) {
					if(!$objectPhpExcel->getSheetByName($this->getOnlySheet)) {
						return $sheetDatas;
					}
					$objectPhpExcel->setActiveSheetIndexByName($this->getOnlySheet);
					$indexed = $this->getOnlySheet;
					$sheetDatas[$indexed] = $objectPhpExcel->getActiveSheet()->toArray(null, true, true, true);
					if ($this->setFirstRecordAsKeys) {
						$sheetDatas[$indexed] = $this->executeArrayLabel($sheetDatas[$indexed]);
					}
					if (!empty($this->getOnlyRecordByIndex)) {
						$sheetDatas[$indexed] = $this->executeGetOnlyRecords($sheetDatas[$indexed], $this->getOnlyRecordByIndex);
					}
					if (!empty($this->leaveRecordByIndex)) {
						$sheetDatas[$indexed] = $this->executeLeaveRecords($sheetDatas[$indexed], $this->leaveRecordByIndex);
					}
					return $sheetDatas[$indexed];
				} else {
					$objectPhpExcel->setActiveSheetIndexByName($sheetName);
					$indexed = $this->setIndexSheetByName==true ? $sheetName : $sheetIndex;
					$sheetDatas[$indexed] = $objectPhpExcel->getActiveSheet()->toArray(null, true, true, true);
					if ($this->setFirstRecordAsKeys) {
						$sheetDatas[$indexed] = $this->executeArrayLabel($sheetDatas[$indexed]);
					}
					if (!empty($this->getOnlyRecordByIndex) && isset($this->getOnlyRecordByIndex[$indexed]) && is_array($this->getOnlyRecordByIndex[$indexed])) {
						$sheetDatas = $this->executeGetOnlyRecords($sheetDatas, $this->getOnlyRecordByIndex[$indexed]);
					}
					if (!empty($this->leaveRecordByIndex) && isset($this->leaveRecordByIndex[$indexed]) && is_array($this->leaveRecordByIndex[$indexed])) {
						$sheetDatas[$indexed] = $this->executeLeaveRecords($sheetDatas[$indexed], $this->leaveRecordByIndex[$indexed]);
					}
				}
			}
		} else {
			$sheetDatas = $objectPhpExcel->getActiveSheet()->toArray(null, true, true, true);
			if ($this->setFirstRecordAsKeys) {
				$sheetDatas = $this->executeArrayLabel($sheetDatas);
			}
			if (!empty($this->getOnlyRecordByIndex)) {
				$sheetDatas = $this->executeGetOnlyRecords($sheetDatas, $this->getOnlyRecordByIndex);
			}
			if (!empty($this->leaveRecordByIndex)) {
				$sheetDatas = $this->executeLeaveRecords($sheetDatas, $this->leaveRecordByIndex);
			}
		}

		return $sheetDatas;
	}

	/**
	 * (non-PHPdoc)
	 * @see \yii\base\Widget::run()
	 */
	public function run()
	{
		if ($this->mode == 'export')
		{
        	    	$sheet = new Spreadsheet();
        	    	
        	    	if (!isset($this->models))
        	    		throw new InvalidConfigException('Config models must be set');
        	    	
        	    	if (isset($this->properties))
        	    	{
        	    		$this->properties($sheet, $this->properties);
        	    	}
        	    	
        	    	if ($this->isMultipleSheet) {
        	    		$index = 0;
        	    		$worksheet = [];
        	    		foreach ($this->models as $title => $models) {
        	    			$sheet->createSheet($index);
        	    			$sheet->getSheet($index)->setTitle($title);
        	    			$worksheet[$index] = $sheet->getSheet($index);
        	    			$columns = isset($this->columns[$title]) ? $this->columns[$title] : [];
        	    			$headers = isset($this->headers[$title]) ? $this->headers[$title] : [];
        	    			$this->executeColumns($models,$worksheet[$index], $this->populateColumns($columns), $headers);
        	    			$index++;
        	    		}
        	    	} else {
        	    		$worksheet = $sheet->getActiveSheet();
        	    		$this->executeColumns($this->models,$worksheet, isset($this->columns) ? $this->populateColumns($this->columns) : [], isset($this->headers) ? $this->headers : []);
        	    	}
        	    	
        	    	if ($this->asAttachment) {
        	    		$this->setHeaders();
        	    	}
	       	$this->writeFile($sheet);
	       	$sheet->disconnectWorksheets();
	       	unset($sheet);
		} 
		elseif ($this->mode == 'import') 
		{
			if (is_array($this->fileName)) {
				$datas = [];
				foreach ($this->fileName as $key => $filename) {
					$datas[$key] = $this->readFile($filename);
				}
				return $datas;
			} else {
				return $this->readFile($this->fileName);
			}
		}
	}

	/**
	 * Exporting data into an excel file.
	 *
	 * ~~~
	 *
	 * \moonland\phpexcel\Excel::export([
	 * 		'models' => $allModels,
	 * 		'columns' => ['column1','column2','column3'],
	 * 		//without header working, because the header will be get label from attribute label.
	 * 		'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
	 * ]);
	 *
	 * ~~~
	 *
	 * New Feature for exporting data, you can use this if you familiar yii gridview.
	 * That is same with gridview data column.
	 * Columns in array mode valid params are 'attribute', 'header', 'format', 'value', and footer (TODO).
	 * Columns in string mode valid layout are 'attribute:format:header:footer(TODO)'.
	 *
	 * ~~~
	 *
	 * \moonland\phpexcel\Excel::export([
	 *  	'models' => Post::find()->all(),
	 *     	'columns' => [
	 *     		'author.name:text:Author Name',
	 *     		[
	 *     				'attribute' => 'content',
	 *     				'header' => 'Content Post',
	 *     				'format' => 'text',
	 *     				'value' => function($model) {
	 *     					return ExampleClass::removeText('example', $model->content);
	 *     				},
	 *     		],
	 *     		'like_it:text:Reader like this content',
	 *     		'created_at:datetime',
	 *     		[
	 *     				'attribute' => 'updated_at',
	 *     				'format' => 'date',
	 *     		],
	 *     	],
	 *     	'headers' => [
	 *     		'created_at' => 'Date Created Content',
	 * 		],
	 * ]);
	 *
	 * ~~~
	 *
	 * @param array $config
	 * @return string
	 */
	public static function export($config=[])
	{
		$config = ArrayHelper::merge(['mode' => 'export'], $config);
		return self::widget($config);
	}

	/**
	 * Import file excel and return into an array.
	 *
	 * ~~~
	 *
	 * $data = \moonland\phpexcel\Excel::import($fileName, ['setFirstRecordAsKeys' => true]);
	 *
	 * ~~~
	 *
	 * @param string!array $fileName to load.
	 * @param array $config is a more configuration.
	 * @return string
	 */
	public static function import($fileName, $config=[])
	{
		$config = ArrayHelper::merge(['mode' => 'import', 'fileName' => $fileName, 'asArray' => true], $config);
		return self::widget($config);
	}

	/**
	 * @param array $config
	 * @return string
	 */
	public static function widget($config = [])
	{
		if ((isset($config['mode']) and $config['mode'] == 'import') && !isset($config['asArray'])) {
			$config['asArray'] = true;
		}

		if (isset($config['asArray']) && $config['asArray']==true)
		{
	        $config['class'] = get_called_class();
	        $widget = \Yii::createObject($config);
	        return $widget->run();
		} else {
			return parent::widget($config);
		}
	}
}

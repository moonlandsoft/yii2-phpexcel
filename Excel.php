<?php

namespace moonland\phpexcel;

/**
 * Excel Widget for generate Excel File.
 * 
 * Usage
 * -----
 * 
 * Once the extension is installed, simply use it in your code by  :
 * 
 * ```php
 * <?php 
 *  \moonland\phpexcel\Excel::widget([
 * 		'models' => $allModels,
 * 		'columns' => ['column1','column2','column3'],
 * 		//without header working, because the header will be get label from attribute label.
 * 		'header' => ['column1' => 'Header Column 1','column2' => 'Header Column 2', 'column3' => 'Header Column 3'],
 * 	]); ?>
 * ```
 * 
 * @author Moh Khoirul Anam <moh.khoirul.anaam@gmail.com>
 * @copyright 2014
 * @since 1
 */
class Excel extends \yii\base\Widget
{
	private $_sheet;
	private static $_sheets;
	private $_activeSheet;
	private static $_activeSheets;
	
	public $models;
	public $columns;
	public $headers;
	public $fileName;
	public $savePath;
	public $format = 'Excel2007';
	public $setFirstTitle = true;
	public $asAttachment = true;
	
	public function getSheet()
	{
		if (!isset($this->_sheet)) {
			$this->_sheet = new \PHPExcel();
		}
		return $this->_sheet;
	}
	
	public static function sheet()
	{
		if (!isset(self::$_sheets)) {
			self::$_sheets = new \PHPExcel();
		}
		return self::$_sheets;
	}
	
	public function getActiveSheet()
	{
		if (!isset($this->_activeSheet)) {
			$this->_activeSheet = $this->sheet()->getActiveSheet();
		}
		return $this->_activeSheet;
	}
	
	public static function activeSheet()
	{
		if (!isset(self::$_activeSheets)) {
			self::$_activeSheets = self::sheet()->getActiveSheet();
		}
		return self::$_activeSheets;
	}
	
	/**
	 * Setting data from models
	 */
	public function executeColumns(&$activeSheet = null)
	{
		if ($activeSheet == null) {
			$activeSheet = $this->activeSheet;
		}
		$columns = $this->columns;
		$hasHeader = false;
		$row = 1;
		$char = 26;
		foreach ($this->models as $model) {
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
					if (isset($this->headers[$column])) {
						$activeSheet->setCellValue($col.$row,$this->headers[$column]);
					} else {
						$activeSheet->setCellValue($col.$row,$model->getAttributeLabel($column));
					}
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
				$activeSheet->setCellValue($col.$row,$model->{$column});
				$colnum++;
			}
			$row++;
		}
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
		$fileName = 'exports.xls';
		if (isset($this->fileName)) {
			$fileName = $this->fileName;
			if (strpos($fileName, '.xls') === false)
				$fileName .= '.xls';
		}
		return $fileName;
	}
	
	/**
	 * saving the xls file to download or to path
	 */
	public function save($sheet)
	{
		$objectwriter = \PHPExcel_IOFactory::createWriter($sheet, $this->format);
		$path = 'php://output';
		if (isset($this->savePath) && $this->savePath != null) {
			$path = $this->savePath . '/' . $this->getFileName();
		}
		$objectwriter->save($path);
		exit();
	}
	
    public function run()
    {
    	$sheet = new \PHPExcel();
    	$activeSheet = $sheet->getActiveSheet();
        $this->executeColumns($activeSheet);
        if ($this->asAttachment) {
        	$this->setHeaders();
        }
        $this->save($sheet);
    }
}

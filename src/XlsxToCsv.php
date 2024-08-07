<?php

namespace lovefc;

/*
 * @Author       : lovefc
 * @Date         : 2023-11-25 12:38:43
 * @LastEditors  : lovefc
 * @LastEditTime : 2024-07-13 16:13:19
 * @Description  : 
 * 
 * Copyright (c) 2023 by lovefc, All Rights Reserved. 
 */

class XlsxToCsv
{

	private $workDir;

	private $workFile;

	private $outputDir;

	private $outputFile;

	private $writeNum;

	private $callBack;

	private $type;

	private $reserves;

	private $showlog;

	private $autoWrite;

	private $autoDeleSourceFile;
	
	private $divide; // 是否按照薄来划分

	const TYPE_STRING = 0x01;    // 字符串

	const TYPE_INT = 0x02;       // 整型

	const TYPE_DOUBLE = 0x04;    // 浮点型

	const TYPE_TIMESTAMP = 0x08; // 时间戳，可以将 xlsx 文件中的格式化时间字符转为时间戳	

	// 构造函数
	public function __construct($config = [])
	{
		if ($this->existsClass() === false) {
			die('Xlswriter is not installed.');
		}
		$path = isset($config['path']) ? realpath($config['path']) : '';
		$this->writeNum = isset($config['writeNum']) ? intval($config['writeNum']) : 5000;
		$callback = function ($row, $sheetName) {
			$text = implode(',', $row);
			return $text;
		};
		$this->callBack = $config['callBack'] ?? $callback;
		$this->type = $config['type'] ?? [];
		$this->reserves = $config['reserves'] ?? [];
		$this->showlog = $config['showLog'] ?? true;
		$this->outputDir = $config['output'] ?? '';
		$this->outputFile = $config['outputFile'] ?? '';
		$this->autoWrite = $config['auto'] ?? true;
		$this->autoDeleSourceFile = $config['autoDeleSourceFile'] ?? false;
		$this->divide = $config['divide'] ?? false;
		if (is_dir($path)) {
			$this->workDir = $path;
		}
		if (is_file($path)) {
			$this->workDir = dirname($path);
			$this->workFile = basename($path);
		}
		if (empty($this->workDir)) {
			die('The directory is wrong.');
		}
		if ($this->containsChinese($this->workDir)) {
			$this->workDir = mb_convert_encoding($this->workDir, "GBK", "UTF-8");
		}
		if ($this->containsChinese($this->workFile)) {
			$this->workFile = mb_convert_encoding($this->workFile, "GBK", "UTF-8");
		}
		if (!empty($this->outputDir) && (!is_dir($this->outputDir))) {
			die('The output directory is wrong.');
		}
	}

	// 检测类是否存在
	public function existsClass()
	{
		if (class_exists('\Vtiful\Kernel\Excel', false)) {
			return true;
		} else {
			return false;
		}
	}

	// 创建目录
	private function createDirectory($directoryPath)
	{
		if (!is_dir($directoryPath)) {
			if (mkdir($directoryPath, 0777, true)) {
				return true;
			} else {
				return false;
			}
		} else {
			return true;
		}
	}

	// 取出所有的文件
	public function scanFile($path, $desiredExtension = '')
	{
		global $result;
		if (!is_dir($path)) {
			die('The directory is wrong.');
		}
		$files = scandir($path);
		foreach ($files as $file) {
			if ($file != '.' && $file != '..') {
				if (is_dir($path . '/' . $file)) {
					$this->scanFile($path . '/' . $file, $desiredExtension);
				} else {
					if (!empty($desiredExtension)) {
						$fileParts = pathinfo($file);
						if (isset($fileParts['extension']) && $fileParts['extension'] === $desiredExtension) {
							$result[] = $path . '/' . $file;
						}
					} else {
						$result[] = $path . '/' . $file;
					}
				}
			}
		}
		return $result;
	}

	// 运行
	public function run()
	{
		if (!empty($this->workFile)) {
			$this->saveExcelToCsv($this->workFile);
		} else {
			$res = $this->scanFile($this->workDir, 'xlsx');
			if (!$res) {
				die('There are no files to process in the current directory.');
			}
			foreach ($res as $v) {
				$pathInfo = pathinfo($v);
				$chree = str_replace($this->workDir, '', $pathInfo['dirname']);
				$name = $pathInfo['basename'];
				$file =  $chree . '/' . $pathInfo['basename'];
				if ($this->showlog === true) {
					echo 'File Name:' . $file . PHP_EOL;
				}
				if (substr($name, 0, 2) != '~$') {
					$filename = $this->checkFilename($file);
					if ($filename) {
						$this->saveExcelToCsv($filename);
						if ($this->autoDeleSourceFile) {
							$this->deleFile($file);
						}
					}
				}
			}
		}
	}

	// 检测文件名
	public function checkFilename($filename)
	{
		if ($this->containsChinese($filename)) {
			$filename = mb_convert_encoding($filename, "GBK", "UTF-8");
		}
		$path = $this->workDir . $filename;
		$filename = trim($filename, '/');
		if (is_file($path)) {
			return $filename;
		}
		if ($this->showlog === true) {
			echo 'File:' . $path . '-File does not exist,skipping automatically.' . PHP_EOL;
		}
		return false;
	}

	// 删除文件
	public function deleFile($filename)
	{
		if ($this->containsChinese($filename)) {
			$filename = mb_convert_encoding($filename, "GBK", "UTF-8");
		}
		$path = $this->workDir . $filename;
		if (is_file($path)) {
			unlink($path);
		}
	}

	// 搜索列名
	public function search($myArray)
	{
		$arr = [];
		$key = $this->reserves;
		if (empty($key) || !is_array($key)) {
			return [];
		}
		foreach ($key as $k => $v) {
			$v = trim($v);
			$value = array_search($v, $myArray);
			if (!isset($arr[$k]) || (!empty($value) || ($value === 0))) {
				$arr[$k] = $value;
			}
		}
		$newArray =  array_filter($arr, function ($value) {
			return $value !== null && $value !== '' && $value !== false;
		});;
		return $newArray;
	}

	// 保存csv
	private function saveCsv($newcsv, $text)
	{
		if ($this->autoWrite === true) {
			file_put_contents($newcsv, $text, FILE_APPEND);
		}
	}

	// 识别中文
	private function containsChinese($string)
	{
		return preg_match('/[\x{4e00}-\x{9fa5}]/u', $string);
	}

	// 读取保存
	public function saveExcelToCsv($filename, $newcsv = '')
	{
		try {
			$excel = new \Vtiful\Kernel\Excel(['path' => $this->workDir]);
			$sheetList = $excel->openFile($filename)->sheetList();
			$sheetCount = count($sheetList);
			$output = !empty($this->outputDir) ? $this->outputDir : $this->workDir;
			$pathInfo = pathinfo($filename);
			$dirname  = $pathInfo['dirname'];
			$filename  = $pathInfo['filename'];
			$name = pathinfo($filename, PATHINFO_FILENAME);
			$directory = ($dirname!= '.') ? $output . '/' . $dirname : $output;
			if($this->divide == false){
			    $newcsv = $directory . '/' . $filename . '.csv';
			}else{
				$newcsv = $directory;
			}
			if ($this->containsChinese($newcsv)) {
				$newcsv = mb_convert_encoding($newcsv, "UTF-8", "GBK");
			}
			if (!empty($this->outputFile)) {
				$newcsv2 = $this->outputFile;
				if (is_dir(dirname($newcsv2))) {
					$newcsv = $newcsv2;
				}
			}	
			if (($this->createDirectory($directory) === false)) {
				die('Directory cannot be created or already exists.');
			}
			$wcount = $this->writeNum;
			foreach ($sheetList as $sheetName) {
				$sheetData = $excel->openSheet($sheetName);
				if (!empty($this->type)) {
					$sheetData->setType($this->type);
				}
				if ($this->showlog === true) {
					echo 'Sheet Name:' . $sheetName . PHP_EOL;
				}
				$i = $i2 = 0;
				$sheet_count = 1;
				$arr = $s_arr =  [];
				if($this->divide == true && is_dir($newcsv)){
				    $newcsv = $newcsv.$sheet_count.'.csv';
				}
				while (true) {
					$row = $sheetData->nextRow();
					if ($row === NULL) {
						$text = implode("", $arr);
						$this->saveCsv($newcsv, $text);
						$arr = [];
						break;
					}
					$row2 = [];
					if ($i == 0) {
						$s_arr = $this->search($row);
						$i = 1;
					}
					if (!empty($s_arr)) {
						$s_arr = array_values($s_arr);
						foreach ($s_arr as $v) {
							$row2[] = $row[$v] ?? '';
						}
					} else {
						$row2 = $row;
					}
					$txt = false;
					if (is_callable($this->callBack)) {
						if ($row2 !== []) {
							$txt = call_user_func($this->callBack, $row2, $sheetName);
						}
					}
					if ($txt) {
						$arr[] = $txt . PHP_EOL;
					}
					$i2 = count($arr);
					if (($i2 != 1) && ($i2 % $wcount == 1)) {
						$text = implode("", $arr);
						$this->saveCsv($newcsv, $text);
						$arr = [];
					}
				}
				$sheet_count++;
			}
		} catch (Exception $e) {
			var_dump($e->getMessage());
		}
	}
}

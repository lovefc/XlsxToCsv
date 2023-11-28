<?php

namespace lovefc;

/*
 * @Author       : lovefc
 * @Date         : 2023-11-25 12:38:43
 * @LastEditors  : lovefc
 * @LastEditTime : 2023-11-28 17:18:02
 * @Description  : 
 * 
 * Copyright (c) 2023 by lovefc, All Rights Reserved. 
 */

class XlsxToCsv
{

	private $workDir;

	private $workFile;

	private $outputDir;

	private $writeNum;

	private $callBack;

	private $type;

	private $reserves;

	private $showlog;

	private $autoWrite;

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
		$this->autoWrite = $config['auto'] ?? true;
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
				$file =  basename($v);
				$this->saveExcelToCsv($file);
			}
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

	// 读取保存
	public function saveExcelToCsv($filename, $newcsv = '')
	{
		try {
			$excel = new \Vtiful\Kernel\Excel(['path' => $this->workDir]);
			$sheetList = $excel->openFile($filename)->sheetList();
			$output = !empty($this->outputDir) ? $this->outputDir : $this->workDir;
			$directory = $output . '/' . pathinfo($filename, PATHINFO_FILENAME);
			if ($this->createDirectory($directory) === false) {
				die('Directory cannot be created or already exists.');
			}
			if ($this->showlog === true) {
				echo 'File Name:' . $filename . PHP_EOL;
			}
			$wcount = $this->writeNum;
			foreach ($sheetList as $sheetName) {
				$newcsv = $directory . '/' . $sheetName . '.csv';
				$sheetData = $excel->openSheet($sheetName);
				if (!empty($this->type)) {
					$sheetData->setType($this->type);
				}
				if ($this->showlog === true) {
					echo 'Sheet Name:' . $sheetName . PHP_EOL;
				}
				$i = $i2 = 0;
				$arr = $s_arr =  [];
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
			}
		} catch (Exception $e) {
			var_dump($e->getMessage());
		}
	}
}
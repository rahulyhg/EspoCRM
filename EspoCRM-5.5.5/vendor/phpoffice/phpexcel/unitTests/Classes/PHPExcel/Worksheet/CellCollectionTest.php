<?php

class CellCollectionTest extends PHPUnit_Framework_TestCase
{

	public function setUp()
	{
		if (!defined('PHPEXCEL_ROOT'))
		{
			define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
		}
		require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}


	public function testCacheLastCell()
	{
		$methods = \PhpOffice\PhpSpreadsheet\Collection\CellsFactory::getCacheStorageMethods();
		foreach ($methods as $method) {
			\PhpOffice\PhpSpreadsheet\Collection\CellsFactory::initialize($method);
			$workbook = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
			$cells = array('A1', 'A2');
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setCellValue('A1', 1);
			$worksheet->setCellValue('A2', 2);
			$this->assertEquals($cells, $worksheet->getCellCollection(), "Cache method \"$method\".");
			\PhpOffice\PhpSpreadsheet\Collection\CellsFactory::finalize();
		}
	}

}

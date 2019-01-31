<?php

class WorksheetRowTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRow;

	public function setUp()
	{
		if (!defined('PHPEXCEL_ROOT')) {
			define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
		}
		require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
        
        $this->mockWorksheet = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();
        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestColumn')
                 ->will($this->returnValue('E'));
    }


	public function testInstantiateRowDefault()
	{
        $row = new \PhpOffice\PhpSpreadsheet\Worksheet\Row($this->mockWorksheet);
        $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Row', $row);
        $rowIndex = $row->getRowIndex();
        $this->assertEquals(1, $rowIndex);
	}

	public function testInstantiateRowSpecified()
	{
        $row = new \PhpOffice\PhpSpreadsheet\Worksheet\Row($this->mockWorksheet, 5);
        $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Row', $row);
        $rowIndex = $row->getRowIndex();
        $this->assertEquals(5, $rowIndex);
	}

	public function testGetCellIterator()
	{
        $row = new \PhpOffice\PhpSpreadsheet\Worksheet\Row($this->mockWorksheet);
        $cellIterator = $row->getCellIterator();
        $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator', $cellIterator);
	}
}

<?php

class WorksheetColumnTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockColumn;

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
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
    }


	public function testInstantiateColumnDefault()
	{
        $column = new \PhpOffice\PhpSpreadsheet\Worksheet\Column($this->mockWorksheet);
        $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Column', $column);
        $columnIndex = $column->getColumnIndex();
        $this->assertEquals('A', $columnIndex);
	}

	public function testInstantiateColumnSpecified()
	{
        $column = new \PhpOffice\PhpSpreadsheet\Worksheet\Column($this->mockWorksheet, 'E');
        $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Column', $column);
        $columnIndex = $column->getColumnIndex();
        $this->assertEquals('E', $columnIndex);
	}

	public function testGetCellIterator()
	{
        $column = new \PhpOffice\PhpSpreadsheet\Worksheet\Column($this->mockWorksheet);
        $cellIterator = $column->getCellIterator();
        $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator', $cellIterator);
	}
}

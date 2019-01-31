<?php

class RowCellIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRowCell;

	public function setUp()
	{
		if (!defined('PHPEXCEL_ROOT')) {
			define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
		}
		require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
        
        $this->mockCell = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Cell\Cell')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestColumn')
                 ->will($this->returnValue('E'));
        $this->mockWorksheet->expects($this->any())
                 ->method('getCellByColumnAndRow')
                 ->will($this->returnValue($this->mockCell));
    }


	public function testIteratorFullRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator($this->mockWorksheet);
        $RowCellIndexResult = 'A';
        $this->assertEquals($RowCellIndexResult, $iterator->key());
        
        foreach($iterator as $key => $RowCell) {
            $this->assertEquals($RowCellIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Cell\Cell', $RowCell);
        }
	}

	public function testIteratorStartEndRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $RowCellIndexResult = 'B';
        $this->assertEquals($RowCellIndexResult, $iterator->key());
        
        foreach($iterator as $key => $RowCell) {
            $this->assertEquals($RowCellIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Cell\Cell', $RowCell);
        }
	}

	public function testIteratorSeekAndPrev()
	{
        $ranges = range('A','E');
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $RowCellIndexResult = 'D';
        $iterator->seek('D');
        $this->assertEquals($RowCellIndexResult, $iterator->key());

        for($i = 1; $i < array_search($RowCellIndexResult, $ranges); $i++) {
            $iterator->prev();
            $expectedResult = $ranges[array_search($RowCellIndexResult, $ranges) - $i];
            $this->assertEquals($expectedResult, $iterator->key());
        }
	}

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $iterator->seek(1);
    }

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $iterator->prev();
    }

}

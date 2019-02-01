<?php

class ColumnCellIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockColumnCell;

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
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
        $this->mockWorksheet->expects($this->any())
                 ->method('getCellByColumnAndRow')
                 ->will($this->returnValue($this->mockCell));
    }


	public function testIteratorFullRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->mockWorksheet, 'A');
        $ColumnCellIndexResult = 1;
        $this->assertEquals($ColumnCellIndexResult, $iterator->key());
        
        foreach($iterator as $key => $ColumnCell) {
            $this->assertEquals($ColumnCellIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Cell\Cell', $ColumnCell);
        }
	}

	public function testIteratorStartEndRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->mockWorksheet, 'A', 2, 4);
        $ColumnCellIndexResult = 2;
        $this->assertEquals($ColumnCellIndexResult, $iterator->key());
        
        foreach($iterator as $key => $ColumnCell) {
            $this->assertEquals($ColumnCellIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Cell\Cell', $ColumnCell);
        }
	}

	public function testIteratorSeekAndPrev()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->mockWorksheet, 'A', 2, 4);
        $columnIndexResult = 4;
        $iterator->seek(4);
        $this->assertEquals($columnIndexResult, $iterator->key());

        for($i = 1; $i < $columnIndexResult-1; $i++) {
            $iterator->prev();
            $this->assertEquals($columnIndexResult - $i, $iterator->key());
        }
	}

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->mockWorksheet, 'A', 2, 4);
        $iterator->seek(1);
    }

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->mockWorksheet, 'A', 2, 4);
        $iterator->prev();
    }

}

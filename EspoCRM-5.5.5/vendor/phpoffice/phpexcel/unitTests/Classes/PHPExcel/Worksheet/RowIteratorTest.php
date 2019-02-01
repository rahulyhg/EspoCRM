<?php

class RowIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRow;

	public function setUp()
	{
		if (!defined('PHPEXCEL_ROOT')) {
			define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
		}
		require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
        
        $this->mockRow = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Row')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
        $this->mockWorksheet->expects($this->any())
                 ->method('current')
                 ->will($this->returnValue($this->mockRow));
    }


	public function testIteratorFullRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator($this->mockWorksheet);
        $rowIndexResult = 1;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Row', $row);
        }
	}

	public function testIteratorStartEndRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $rowIndexResult = 2;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Row', $row);
        }
	}

	public function testIteratorSeekAndPrev()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
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
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $iterator->seek(1);
    }

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $iterator->prev();
    }

}

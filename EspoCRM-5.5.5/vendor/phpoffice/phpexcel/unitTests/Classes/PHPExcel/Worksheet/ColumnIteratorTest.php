<?php

class ColumnIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockColumn;

	public function setUp()
	{
		if (!defined('PHPEXCEL_ROOT')) {
			define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
		}
		require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
        
        $this->mockColumn = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Column')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestColumn')
                 ->will($this->returnValue('E'));
        $this->mockWorksheet->expects($this->any())
                 ->method('current')
                 ->will($this->returnValue($this->mockColumn));
    }


	public function testIteratorFullRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator($this->mockWorksheet);
        $columnIndexResult = 'A';
        $this->assertEquals($columnIndexResult, $iterator->key());
        
        foreach($iterator as $key => $column) {
            $this->assertEquals($columnIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Column', $column);
        }
	}

	public function testIteratorStartEndRange()
	{
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator($this->mockWorksheet, 'B', 'D');
        $columnIndexResult = 'B';
        $this->assertEquals($columnIndexResult, $iterator->key());
        
        foreach($iterator as $key => $column) {
            $this->assertEquals($columnIndexResult++, $key);
            $this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\Column', $column);
        }
	}

	public function testIteratorSeekAndPrev()
	{
        $ranges = range('A','E');
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator($this->mockWorksheet, 'B', 'D');
        $columnIndexResult = 'D';
        $iterator->seek('D');
        $this->assertEquals($columnIndexResult, $iterator->key());

        for($i = 1; $i < array_search($columnIndexResult, $ranges); $i++) {
            $iterator->prev();
            $expectedResult = $ranges[array_search($columnIndexResult, $ranges) - $i];
            $this->assertEquals($expectedResult, $iterator->key());
        }
	}

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator($this->mockWorksheet, 'B', 'D');
        $iterator->seek('A');
    }

    /**
     * @expectedException \PhpOffice\PhpSpreadsheet\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator($this->mockWorksheet, 'B', 'D');
        $iterator->prev();
    }

}

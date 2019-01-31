<?php


class XEEValidatorTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

    /**
     * @dataProvider providerInvalidXML
     * @expectedException \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
	public function testInvalidXML($filename)
	{
        $reader = $this->getMockForAbstractClass('\PhpOffice\PhpSpreadsheet\Reader\BaseReader');
        $expectedResult = 'FAILURE: Should throw an Exception rather than return a value';
		$result = $reader->securityScanFile($filename);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerInvalidXML()
    {
        $tests = [];
        foreach(glob('rawTestData/Reader/XEETestInvalid*.xml') as $file) {
            $tests[] = [realpath($file), true];
        }
        return $tests;
	}

    /**
     * @dataProvider providerValidXML
     */
	public function testValidXML($filename, $expectedResult)
	{
        $reader = $this->getMockForAbstractClass('\PhpOffice\PhpSpreadsheet\Reader\BaseReader');
		$result = $reader->securityScanFile($filename);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerValidXML()
    {
        $tests = [];
        foreach(glob('rawTestData/Reader/XEETestValid*.xml') as $file) {
            $tests[] = [realpath($file), file_get_contents($file)];
        }
        return $tests;
	}

}

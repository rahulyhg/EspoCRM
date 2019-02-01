<?php


class AutoloaderTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT'))
        {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
    }

    public function testAutoloaderNon\PhpOffice\PhpSpreadsheet\SpreadsheetClass()
    {
        $className = 'InvalidClass';

        $result = \PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    public function testAutoloaderInvalid\PhpOffice\PhpSpreadsheet\SpreadsheetClass()
    {
        $className = '\PhpOffice\PhpSpreadsheet\Spreadsheet_Invalid_Class';

        $result = \PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    public function testAutoloadValid\PhpOffice\PhpSpreadsheet\SpreadsheetClass()
    {
        $className = '\PhpOffice\PhpSpreadsheet\IOFactory';

        $result = \PhpOffice\PhpSpreadsheet\Spreadsheet_Autoloader::Load($className);
        //    Check that class has been loaded
        $this->assertTrue(class_exists($className));
    }

    public function testAutoloadInstantiateSuccess()
    {
        $result = new \PhpOffice\PhpSpreadsheet\Calculation\Category(1,2,3);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result,'\PhpOffice\PhpSpreadsheet\Calculation\Category'));
    }

}
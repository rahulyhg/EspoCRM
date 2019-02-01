<?php

require_once 'testDataFileIterator.php';

class CalculationTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');

        \PhpOffice\PhpSpreadsheet\Calculation\Functions::setCompatibilityMode(\PhpOffice\PhpSpreadsheet\Calculation\Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerBinaryComparisonOperation
     */
    public function testBinaryComparisonOperation($formula, $expectedResultExcel, $expectedResultOpenOffice)
    {
        \PhpOffice\PhpSpreadsheet\Calculation\Functions::setCompatibilityMode(\PhpOffice\PhpSpreadsheet\Calculation\Functions::COMPATIBILITY_EXCEL);
        $resultExcel = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance()->_calculateFormulaValue($formula);
        $this->assertEquals($expectedResultExcel, $resultExcel, 'should be Excel compatible');

        \PhpOffice\PhpSpreadsheet\Calculation\Functions::setCompatibilityMode(\PhpOffice\PhpSpreadsheet\Calculation\Functions::COMPATIBILITY_OPENOFFICE);
        $resultOpenOffice = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance()->_calculateFormulaValue($formula);
        $this->assertEquals($expectedResultOpenOffice, $resultOpenOffice, 'should be OpenOffice compatible');
    }

    public function providerBinaryComparisonOperation()
    {
        return new testDataFileIterator('rawTestData/CalculationBinaryComparisonOperation.data');
    }

}

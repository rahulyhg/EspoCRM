<?php


class DataSeriesValuesTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT'))
        {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testSetDataType()
	{
		$dataTypeValues = array(
			'Number',
			'String'
		);

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;

		foreach($dataTypeValues as $dataTypeValue) {
			$result = $testInstance->setDataType($dataTypeValue);
			$this->assertTrue($result instanceof \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues);
		}
	}

	public function testSetInvalidDataTypeThrowsException()
	{
		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;

		try {
			$result = $testInstance->setDataType('BOOLEAN');
		} catch (Exception $e) {
			$this->assertEquals($e->getMessage(), 'Invalid datatype for chart data series values');
			return;
		}
		$this->fail('An expected exception has not been raised.');
	}

	public function testGetDataType()
	{
		$dataTypeValue = 'String';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
		$setValue = $testInstance->setDataType($dataTypeValue);

		$result = $testInstance->getDataType();
		$this->assertEquals($dataTypeValue,$result);
	}

}

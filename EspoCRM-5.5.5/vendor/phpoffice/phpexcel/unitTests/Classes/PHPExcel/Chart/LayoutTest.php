<?php


class LayoutTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT'))
        {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testSetLayoutTarget()
	{
		$LayoutTargetValue = 'String';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Layout;

		$result = $testInstance->setLayoutTarget($LayoutTargetValue);
		$this->assertTrue($result instanceof \PhpOffice\PhpSpreadsheet\Chart\Layout);
	}

	public function testGetLayoutTarget()
	{
		$LayoutTargetValue = 'String';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Layout;
		$setValue = $testInstance->setLayoutTarget($LayoutTargetValue);

		$result = $testInstance->getLayoutTarget();
		$this->assertEquals($LayoutTargetValue,$result);
	}

}

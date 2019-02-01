<?php


class LegendTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT'))
        {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testSetPosition()
	{
		$positionValues = array(
			\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_RIGHT,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_LEFT,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_TOP,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_BOTTOM,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_TOPRIGHT,
		);

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;

		foreach($positionValues as $positionValue) {
			$result = $testInstance->setPosition($positionValue);
			$this->assertTrue($result);
		}
	}

	public function testSetInvalidPositionReturnsFalse()
	{
		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;

		$result = $testInstance->setPosition('BottomLeft');
		$this->assertFalse($result);
		//	Ensure that value is unchanged
		$result = $testInstance->getPosition();
		$this->assertEquals(\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_RIGHT,$result);
	}

	public function testGetPosition()
	{
		$PositionValue = \PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_BOTTOM;

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;
		$setValue = $testInstance->setPosition($PositionValue);

		$result = $testInstance->getPosition();
		$this->assertEquals($PositionValue,$result);
	}

	public function testSetPositionXL()
	{
		$positionValues = array(
			\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionBottom,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionCorner,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionCustom,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionLeft,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionRight,
			\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionTop,
		);

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;

		foreach($positionValues as $positionValue) {
			$result = $testInstance->setPositionXL($positionValue);
			$this->assertTrue($result);
		}
	}

	public function testSetInvalidXLPositionReturnsFalse()
	{
		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;

		$result = $testInstance->setPositionXL(999);
		$this->assertFalse($result);
		//	Ensure that value is unchanged
		$result = $testInstance->getPositionXL();
		$this->assertEquals(\PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionRight,$result);
	}

	public function testGetPositionXL()
	{
		$PositionValue = \PhpOffice\PhpSpreadsheet\Chart\Legend::xlLegendPositionCorner;

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;
		$setValue = $testInstance->setPositionXL($PositionValue);

		$result = $testInstance->getPositionXL();
		$this->assertEquals($PositionValue,$result);
	}

	public function testSetOverlay()
	{
		$overlayValues = array(
			TRUE,
			FALSE,
		);

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;

		foreach($overlayValues as $overlayValue) {
			$result = $testInstance->setOverlay($overlayValue);
			$this->assertTrue($result);
		}
	}

	public function testSetInvalidOverlayReturnsFalse()
	{
		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;

		$result = $testInstance->setOverlay('INVALID');
		$this->assertFalse($result);

		$result = $testInstance->getOverlay();
		$this->assertFalse($result);
	}

	public function testGetOverlay()
	{
		$OverlayValue = TRUE;

		$testInstance = new \PhpOffice\PhpSpreadsheet\Chart\Legend;
		$setValue = $testInstance->setOverlay($OverlayValue);

		$result = $testInstance->getOverlay();
		$this->assertEquals($OverlayValue,$result);
	}

}

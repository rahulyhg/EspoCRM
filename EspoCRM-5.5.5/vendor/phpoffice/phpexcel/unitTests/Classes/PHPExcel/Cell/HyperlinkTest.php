<?php


class HyperlinkTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT'))
        {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testGetUrl()
	{
		$urlValue = 'http://www.phpexcel.net';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink($urlValue);

		$result = $testInstance->getUrl();
		$this->assertEquals($urlValue,$result);
	}

	public function testSetUrl()
	{
		$initialUrlValue = 'http://www.phpexcel.net';
		$newUrlValue = 'http://github.com/PHPOffice/\PhpOffice\PhpSpreadsheet\Spreadsheet';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink($initialUrlValue);
		$result = $testInstance->setUrl($newUrlValue);
		$this->assertTrue($result instanceof \PhpOffice\PhpSpreadsheet\Cell\Hyperlink);

		$result = $testInstance->getUrl();
		$this->assertEquals($newUrlValue,$result);
	}

	public function testGetTooltip()
	{
		$tooltipValue = '\PhpOffice\PhpSpreadsheet\Spreadsheet Web Site';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink(NULL, $tooltipValue);

		$result = $testInstance->getTooltip();
		$this->assertEquals($tooltipValue,$result);
	}

	public function testSetTooltip()
	{
		$initialTooltipValue = '\PhpOffice\PhpSpreadsheet\Spreadsheet Web Site';
		$newTooltipValue = '\PhpOffice\PhpSpreadsheet\Spreadsheet Repository on Github';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink(NULL, $initialTooltipValue);
		$result = $testInstance->setTooltip($newTooltipValue);
		$this->assertTrue($result instanceof \PhpOffice\PhpSpreadsheet\Cell\Hyperlink);

		$result = $testInstance->getTooltip();
		$this->assertEquals($newTooltipValue,$result);
	}

	public function testIsInternal()
	{
		$initialUrlValue = 'http://www.phpexcel.net';
		$newUrlValue = 'sheet://Worksheet1!A1';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink($initialUrlValue);
		$result = $testInstance->isInternal();
		$this->assertFalse($result);

		$testInstance->setUrl($newUrlValue);
		$result = $testInstance->isInternal();
		$this->assertTrue($result);
	}

	public function testGetHashCode()
	{
		$urlValue = 'http://www.phpexcel.net';
		$tooltipValue = '\PhpOffice\PhpSpreadsheet\Spreadsheet Web Site';
		$initialExpectedHash = 'd84d713aed1dbbc8a7c5af183d6c7dbb';

		$testInstance = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink($urlValue, $tooltipValue);

		$result = $testInstance->getHashCode();
		$this->assertEquals($initialExpectedHash,$result);
	}

}

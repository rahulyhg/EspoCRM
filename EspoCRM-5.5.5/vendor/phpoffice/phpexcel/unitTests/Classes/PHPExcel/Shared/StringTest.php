<?php


require_once 'testDataFileIterator.php';

class StringTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testGetIsMbStringEnabled()
	{
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getIsMbstringEnabled'));
		$this->assertTrue($result);
	}

	public function testGetIsIconvEnabled()
	{
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getIsIconvEnabled'));
		$this->assertTrue($result);
	}

	public function testGetDecimalSeparator()
	{
		$localeconv = localeconv();

		$expectedResult = (!empty($localeconv['decimal_point'])) ? $localeconv['decimal_point'] : ',';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getDecimalSeparator'));
		$this->assertEquals($expectedResult, $result);
	}

	public function testSetDecimalSeparator()
	{
		$expectedResult = ',';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','setDecimalSeparator'),$expectedResult);

		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getDecimalSeparator'));
		$this->assertEquals($expectedResult, $result);
	}

	public function testGetThousandsSeparator()
	{
		$localeconv = localeconv();

		$expectedResult = (!empty($localeconv['thousands_sep'])) ? $localeconv['thousands_sep'] : ',';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getThousandsSeparator'));
		$this->assertEquals($expectedResult, $result);
	}

	public function testSetThousandsSeparator()
	{
		$expectedResult = ' ';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','setThousandsSeparator'),$expectedResult);

		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getThousandsSeparator'));
		$this->assertEquals($expectedResult, $result);
	}

	public function testGetCurrencyCode()
	{
		$localeconv = localeconv();

		$expectedResult = (!empty($localeconv['currency_symbol'])) ? $localeconv['currency_symbol'] : '$';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getCurrencyCode'));
		$this->assertEquals($expectedResult, $result);
	}

	public function testSetCurrencyCode()
	{
		$expectedResult = '£';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','setCurrencyCode'),$expectedResult);

		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\StringHelper','getCurrencyCode'));
		$this->assertEquals($expectedResult, $result);
	}

}

<?php


require_once 'testDataFileIterator.php';

class FontTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testGetAutoSizeMethod()
	{
		$expectedResult = \PhpOffice\PhpSpreadsheet\Shared\Font::AUTOSIZE_METHOD_APPROX;

		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\Font','getAutoSizeMethod'));
		$this->assertEquals($expectedResult, $result);
	}

	public function testSetAutoSizeMethod()
	{
		$autosizeMethodValues = array(
			\PhpOffice\PhpSpreadsheet\Shared\Font::AUTOSIZE_METHOD_EXACT,
			\PhpOffice\PhpSpreadsheet\Shared\Font::AUTOSIZE_METHOD_APPROX,
		);

		foreach($autosizeMethodValues as $autosizeMethodValue) {
			$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\Font','setAutoSizeMethod'),$autosizeMethodValue);
			$this->assertTrue($result);
		}
	}

    public function testSetAutoSizeMethodWithInvalidValue()
	{
		$unsupportedAutosizeMethod = 'guess';

		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\Font','setAutoSizeMethod'),$unsupportedAutosizeMethod);
		$this->assertFalse($result);
	}

    /**
     * @dataProvider providerFontSizeToPixels
     */
	public function testFontSizeToPixels()
	{
		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Font','fontSizeToPixels'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerFontSizeToPixels()
    {
    	return new testDataFileIterator('rawTestData/Shared/FontSizeToPixels.data');
	}

    /**
     * @dataProvider providerInchSizeToPixels
     */
	public function testInchSizeToPixels()
	{
		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Font','inchSizeToPixels'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerInchSizeToPixels()
    {
    	return new testDataFileIterator('rawTestData/Shared/InchSizeToPixels.data');
	}

    /**
     * @dataProvider providerCentimeterSizeToPixels
     */
	public function testCentimeterSizeToPixels()
	{
		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Font','centimeterSizeToPixels'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerCentimeterSizeToPixels()
    {
    	return new testDataFileIterator('rawTestData/Shared/CentimeterSizeToPixels.data');
	}

}

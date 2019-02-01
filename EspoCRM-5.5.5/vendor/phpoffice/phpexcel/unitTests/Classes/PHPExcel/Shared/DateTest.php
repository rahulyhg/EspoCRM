<?php


require_once 'testDataFileIterator.php';

class DateTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');
	}

	public function testSetExcelCalendar()
	{
		$calendarValues = array(
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_MAC_1904,
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900,
		);

		foreach($calendarValues as $calendarValue) {
			$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),$calendarValue);
			$this->assertTrue($result);
		}
	}

    public function testSetExcelCalendarWithInvalidValue()
	{
		$unsupportedCalendar = '2012';
		$result = call_user_func(array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),$unsupportedCalendar);
		$this->assertFalse($result);
	}

    /**
     * @dataProvider providerDateTimeexcelToTimestamp1900
     */
	public function testDateTimeexcelToTimestamp1900()
	{
		$result = call_user_func(
			array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900
		);

		$args = func_get_args();
		$expectedResult = array_pop($args);
		if ($args[0] < 1) {
			$expectedResult += gmmktime(0,0,0);
		}
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','excelToTimestamp'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerDateTimeexcelToTimestamp1900()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimeexcelToTimestamp1900.data');
	}

    /**
     * @dataProvider providerDateTimePHPToExcel1900
     */
	public function testDateTimePHPToExcel1900()
	{
		$result = call_user_func(
			array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900
		);

		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','PHPToExcel'),$args);
		$this->assertEquals($expectedResult, $result, NULL, 1E-5);
	}

    public function providerDateTimePHPToExcel1900()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimePHPToExcel1900.data');
	}

    /**
     * @dataProvider providerDateTimeformattedPHPToExcel1900
     */
	public function testDateTimeformattedPHPToExcel1900()
	{
		$result = call_user_func(
			array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900
		);

		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','formattedPHPToExcel'),$args);
		$this->assertEquals($expectedResult, $result, NULL, 1E-5);
	}

    public function providerDateTimeformattedPHPToExcel1900()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimeformattedPHPToExcel1900.data');
	}

    /**
     * @dataProvider providerDateTimeexcelToTimestamp1904
     */
	public function testDateTimeexcelToTimestamp1904()
	{
		$result = call_user_func(
			array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_MAC_1904
		);

		$args = func_get_args();
		$expectedResult = array_pop($args);
		if ($args[0] < 1) {
			$expectedResult += gmmktime(0,0,0);
		}
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','excelToTimestamp'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerDateTimeexcelToTimestamp1904()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimeexcelToTimestamp1904.data');
	}

    /**
     * @dataProvider providerDateTimePHPToExcel1904
     */
	public function testDateTimePHPToExcel1904()
	{
		$result = call_user_func(
			array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_MAC_1904
		);

		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','PHPToExcel'),$args);
		$this->assertEquals($expectedResult, $result, NULL, 1E-5);
	}

    public function providerDateTimePHPToExcel1904()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimePHPToExcel1904.data');
	}

    /**
     * @dataProvider providerIsDateTimeFormatCode
     */
	public function testIsDateTimeFormatCode()
	{
		$args = func_get_args();
		$expectedResult = array_pop($args);
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','isDateTimeFormatCode'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerIsDateTimeFormatCode()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimeFormatCodes.data');
	}

    /**
     * @dataProvider providerDateTimeexcelToTimestamp1900Timezone
     */
	public function testDateTimeexcelToTimestamp1900Timezone()
	{
		$result = call_user_func(
			array('\PhpOffice\PhpSpreadsheet\Shared\Date','setExcelCalendar'),
			\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900
		);

		$args = func_get_args();
		$expectedResult = array_pop($args);
		if ($args[0] < 1) {
			$expectedResult += gmmktime(0,0,0);
		}
		$result = call_user_func_array(array('\PhpOffice\PhpSpreadsheet\Shared\Date','excelToTimestamp'),$args);
		$this->assertEquals($expectedResult, $result);
	}

    public function providerDateTimeexcelToTimestamp1900Timezone()
    {
    	return new testDataFileIterator('rawTestData/Shared/DateTimeexcelToTimestamp1900Timezone.data');
	}

}

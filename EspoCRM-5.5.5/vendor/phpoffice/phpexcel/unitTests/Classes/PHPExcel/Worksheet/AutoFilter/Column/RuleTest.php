<?php


class RuleTest extends PHPUnit_Framework_TestCase
{
	private $_testAutoFilterRuleObject;

	private $_mockAutoFilterColumnObject;

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPEXCEL_ROOT . '\PhpOffice\PhpSpreadsheet\Spreadsheet/Autoloader.php');

        $this->_mockAutoFilterColumnObject = $this->getMockBuilder('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column')
        	->disableOriginalConstructor()
        	->getMock();

        $this->_mockAutoFilterColumnObject->expects($this->any())
        	->method('testColumnInRange')
        	->will($this->returnValue(3));

		$this->_testAutoFilterRuleObject = new \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule(
			$this->_mockAutoFilterColumnObject
		);
    }

	public function testGetRuleType()
	{
		$result = $this->_testAutoFilterRuleObject->getRuleType();
		$this->assertEquals(\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_FILTER, $result);
	}

	public function testSetRuleType()
	{
		$expectedResult = \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP;

		//	Setters return the instance to implement the fluent interface
		$result = $this->_testAutoFilterRuleObject->setRuleType($expectedResult);
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule', $result);

		$result = $this->_testAutoFilterRuleObject->getRuleType();
		$this->assertEquals($expectedResult, $result);
	}

	public function testSetValue()
	{
		$expectedResult = 100;

		//	Setters return the instance to implement the fluent interface
		$result = $this->_testAutoFilterRuleObject->setValue($expectedResult);
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule', $result);

		$result = $this->_testAutoFilterRuleObject->getValue();
		$this->assertEquals($expectedResult, $result);
	}

	public function testGetOperator()
	{
		$result = $this->_testAutoFilterRuleObject->getOperator();
		$this->assertEquals(\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL, $result);
	}

	public function testSetOperator()
	{
		$expectedResult = \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_LESSTHAN;

		//	Setters return the instance to implement the fluent interface
		$result = $this->_testAutoFilterRuleObject->setOperator($expectedResult);
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule', $result);

		$result = $this->_testAutoFilterRuleObject->getOperator();
		$this->assertEquals($expectedResult, $result);
	}

	public function testSetGrouping()
	{
		$expectedResult = \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MONTH;

		//	Setters return the instance to implement the fluent interface
		$result = $this->_testAutoFilterRuleObject->setGrouping($expectedResult);
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule', $result);

		$result = $this->_testAutoFilterRuleObject->getGrouping();
		$this->assertEquals($expectedResult, $result);
	}

	public function testGetParent()
	{
		$result = $this->_testAutoFilterRuleObject->getParent();
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column', $result);
	}

	public function testSetParent()
	{
		//	Setters return the instance to implement the fluent interface
		$result = $this->_testAutoFilterRuleObject->setParent($this->_mockAutoFilterColumnObject);
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule', $result);
	}

	public function testClone()
	{
		$result = clone $this->_testAutoFilterRuleObject;
		$this->assertInstanceOf('\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule', $result);
	}

}

<?php
/**
 * \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 * Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Style\Style
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Style\Conditional
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Style\Style
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Style\Conditional implements \PhpOffice\PhpSpreadsheet\IComparable
{
	/* Condition types */
	const CONDITION_NONE					= 'none';
	const CONDITION_CELLIS					= 'cellIs';
	const CONDITION_CONTAINSTEXT			= 'containsText';
	const CONDITION_EXPRESSION 				= 'expression';

	/* Operator types */
	const OPERATOR_NONE						= '';
	const OPERATOR_BEGINSWITH				= 'beginsWith';
	const OPERATOR_ENDSWITH					= 'endsWith';
	const OPERATOR_EQUAL					= 'equal';
	const OPERATOR_GREATERTHAN				= 'greaterThan';
	const OPERATOR_GREATERTHANOREQUAL		= 'greaterThanOrEqual';
	const OPERATOR_LESSTHAN					= 'lessThan';
	const OPERATOR_LESSTHANOREQUAL			= 'lessThanOrEqual';
	const OPERATOR_NOTEQUAL					= 'notEqual';
	const OPERATOR_CONTAINSTEXT				= 'containsText';
	const OPERATOR_NOTCONTAINS				= 'notContains';
	const OPERATOR_BETWEEN					= 'between';

	/**
	 * Condition type
	 *
	 * @var int
	 */
	private $_conditionType;

	/**
	 * Operator type
	 *
	 * @var int
	 */
	private $_operatorType;

	/**
	 * Text
	 *
	 * @var string
	 */
	private $_text;

	/**
	 * Condition
	 *
	 * @var string[]
	 */
	private $_condition = array();

	/**
	 * Style
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Style\Style
	 */
	private $_style;

    /**
     * Create a new \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_conditionType		= \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_NONE;
    	$this->_operatorType		= \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_NONE;
    	$this->_text    			= null;
    	$this->_condition			= array();
    	$this->_style				= new \PhpOffice\PhpSpreadsheet\Style\Style(FALSE, TRUE);
    }

    /**
     * Get Condition type
     *
     * @return string
     */
    public function getConditionType() {
    	return $this->_conditionType;
    }

    /**
     * Set Condition type
     *
     * @param string $pValue	\PhpOffice\PhpSpreadsheet\Style\Conditional condition type
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function setConditionType($pValue = \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_NONE) {
    	$this->_conditionType = $pValue;
    	return $this;
    }

    /**
     * Get Operator type
     *
     * @return string
     */
    public function getOperatorType() {
    	return $this->_operatorType;
    }

    /**
     * Set Operator type
     *
     * @param string $pValue	\PhpOffice\PhpSpreadsheet\Style\Conditional operator type
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function setOperatorType($pValue = \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_NONE) {
    	$this->_operatorType = $pValue;
    	return $this;
    }

    /**
     * Get text
     *
     * @return string
     */
    public function getText() {
        return $this->_text;
    }

    /**
     * Set text
     *
     * @param string $value
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function setText($value = null) {
           $this->_text = $value;
           return $this;
    }

    /**
     * Get Condition
     *
     * @deprecated Deprecated, use getConditions instead
     * @return string
     */
    public function getCondition() {
    	if (isset($this->_condition[0])) {
    		return $this->_condition[0];
    	}

    	return '';
    }

    /**
     * Set Condition
     *
     * @deprecated Deprecated, use setConditions instead
     * @param string $pValue	Condition
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function setCondition($pValue = '') {
    	if (!is_array($pValue))
    		$pValue = array($pValue);

    	return $this->setConditions($pValue);
    }

    /**
     * Get Conditions
     *
     * @return string[]
     */
    public function getConditions() {
    	return $this->_condition;
    }

    /**
     * Set Conditions
     *
     * @param string[] $pValue	Condition
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function setConditions($pValue) {
    	if (!is_array($pValue))
    		$pValue = array($pValue);

    	$this->_condition = $pValue;
    	return $this;
    }

    /**
     * Add Condition
     *
     * @param string $pValue	Condition
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function addCondition($pValue = '') {
    	$this->_condition[] = $pValue;
    	return $this;
    }

    /**
     * Get Style
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function getStyle() {
    	return $this->_style;
    }

    /**
     * Set Style
     *
     * @param 	\PhpOffice\PhpSpreadsheet\Style\Style $pValue
     * @throws 	\PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional
     */
    public function setStyle(\PhpOffice\PhpSpreadsheet\Style\Style $pValue = null) {
   		$this->_style = $pValue;
   		return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_conditionType
    		. $this->_operatorType
    		. implode(';', $this->_condition)
    		. $this->_style->getHashCode()
    		. __CLASS__
    	);
    }

	/**
	 * Implement PHP __clone to create a deep clone, not just a shallow copy.
	 */
	public function __clone() {
		$vars = get_object_vars($this);
		foreach ($vars as $key => $value) {
			if (is_object($value)) {
				$this->$key = clone $value;
			} else {
				$this->$key = $value;
			}
		}
	}
}

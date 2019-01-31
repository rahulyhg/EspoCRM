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
 * @package    \PhpOffice\PhpSpreadsheet\Calculation\Calculation
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Calculation\Category
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Calculation\Calculation
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Calculation\Category {
	/* Function categories */
	const CATEGORY_CUBE						= 'Cube';
	const CATEGORY_DATABASE					= 'Database';
	const CATEGORY_DATE_AND_TIME			= 'Date and Time';
	const CATEGORY_ENGINEERING				= 'Engineering';
	const CATEGORY_FINANCIAL				= 'Financial';
	const CATEGORY_INFORMATION				= 'Information';
	const CATEGORY_LOGICAL					= 'Logical';
	const CATEGORY_LOOKUP_AND_REFERENCE		= 'Lookup and Reference';
	const CATEGORY_MATH_AND_TRIG			= 'Math and Trig';
	const CATEGORY_STATISTICAL				= 'Statistical';
	const CATEGORY_TEXT_AND_DATA			= 'Text and Data';

	/**
	 * Category (represented by CATEGORY_*)
	 *
	 * @var string
	 */
	private $_category;

	/**
	 * Excel name
	 *
	 * @var string
	 */
	private $_excelName;

	/**
	 * \PhpOffice\PhpSpreadsheet\Spreadsheet name
	 *
	 * @var string
	 */
	private $_phpExcelName;

    /**
     * Create a new \PhpOffice\PhpSpreadsheet\Calculation\Category
     *
     * @param 	string		$pCategory 		Category (represented by CATEGORY_*)
     * @param 	string		$pExcelName		Excel function name
     * @param 	string		$p\PhpOffice\PhpSpreadsheet\SpreadsheetName	\PhpOffice\PhpSpreadsheet\Spreadsheet function mapping
     * @throws 	\PhpOffice\PhpSpreadsheet\Calculation\Exception
     */
    public function __construct($pCategory = NULL, $pExcelName = NULL, $p\PhpOffice\PhpSpreadsheet\SpreadsheetName = NULL)
    {
    	if (($pCategory !== NULL) && ($pExcelName !== NULL) && ($p\PhpOffice\PhpSpreadsheet\SpreadsheetName !== NULL)) {
    		// Initialise values
    		$this->_category 		= $pCategory;
    		$this->_excelName 		= $pExcelName;
    		$this->_phpExcelName 	= $p\PhpOffice\PhpSpreadsheet\SpreadsheetName;
    	} else {
    		throw new \PhpOffice\PhpSpreadsheet\Calculation\Exception("Invalid parameters passed.");
    	}
    }

    /**
     * Get Category (represented by CATEGORY_*)
     *
     * @return string
     */
    public function getCategory() {
    	return $this->_category;
    }

    /**
     * Set Category (represented by CATEGORY_*)
     *
     * @param 	string		$value
     * @throws 	\PhpOffice\PhpSpreadsheet\Calculation\Exception
     */
    public function setCategory($value = null) {
    	if (!is_null($value)) {
    		$this->_category = $value;
    	} else {
    		throw new \PhpOffice\PhpSpreadsheet\Calculation\Exception("Invalid parameter passed.");
    	}
    }

    /**
     * Get Excel name
     *
     * @return string
     */
    public function getExcelName() {
    	return $this->_excelName;
    }

    /**
     * Set Excel name
     *
     * @param string	$value
     */
    public function setExcelName($value) {
    	$this->_excelName = $value;
    }

    /**
     * Get \PhpOffice\PhpSpreadsheet\Spreadsheet name
     *
     * @return string
     */
    public function get\PhpOffice\PhpSpreadsheet\SpreadsheetName() {
    	return $this->_phpExcelName;
    }

    /**
     * Set \PhpOffice\PhpSpreadsheet\Spreadsheet name
     *
     * @param string	$value
     */
    public function set\PhpOffice\PhpSpreadsheet\SpreadsheetName($value) {
    	$this->_phpExcelName = $value;
    }
}

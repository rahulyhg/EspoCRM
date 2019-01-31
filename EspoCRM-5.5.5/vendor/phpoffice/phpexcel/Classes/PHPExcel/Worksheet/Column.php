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
 * @package    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Worksheet\Column
 *
 * Represents a column in \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, used by \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Worksheet\Column
{
	/**
	 * \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 */
	private $_parent;

	/**
	 * Column index
	 *
	 * @var string
	 */
	private $_columnIndex;

	/**
	 * Create a new column
	 *
	 * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet 	$parent
	 * @param string				$columnIndex
	 */
	public function __construct(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $parent = null, $columnIndex = 'A') {
		// Set parent and column index
		$this->_parent 		= $parent;
		$this->_columnIndex = $columnIndex;
	}

	/**
	 * Destructor
	 */
	public function __destruct() {
		unset($this->_parent);
	}

	/**
	 * Get column index
	 *
	 * @return int
	 */
	public function getColumnIndex() {
		return $this->_columnIndex;
	}

	/**
	 * Get cell iterator
	 *
	 * @param	integer				$startRow	    The row number at which to start iterating
	 * @param	integer				$endRow	        Optionally, the row number at which to stop iterating
	 * @return \PhpOffice\PhpSpreadsheet\Worksheet\CellIterator
	 */
	public function getCellIterator($startRow = 1, $endRow = null) {
		return new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->_parent, $this->_columnIndex, $startRow, $endRow);
	}
}

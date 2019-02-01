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
 * @package	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator
 *
 * Used to iterate columns in a \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator implements Iterator
{
	/**
	 * \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet to iterate
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 */
	private $_subject;

	/**
	 * Current iterator position
	 *
	 * @var int
	 */
	private $_position = 0;

	/**
	 * Start position
	 *
	 * @var int
	 */
	private $_startColumn = 0;


	/**
	 * End position
	 *
	 * @var int
	 */
	private $_endColumn = 0;


	/**
	 * Create a new column iterator
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$subject	The worksheet to iterate over
	 * @param	string				$startColumn	The column address at which to start iterating
	 * @param	string				$endColumn	    Optionally, the column address at which to stop iterating
	 */
	public function __construct(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $subject = null, $startColumn = 'A', $endColumn = null) {
		// Set subject
		$this->_subject = $subject;
		$this->resetEnd($endColumn);
		$this->resetStart($startColumn);
	}

	/**
	 * Destructor
	 */
	public function __destruct() {
		unset($this->_subject);
	}

	/**
	 * (Re)Set the start column and the current column pointer
	 *
	 * @param integer	$startColumn	The column address at which to start iterating
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator
	 */
	public function resetStart($startColumn = 'A') {
        $startColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($startColumn) - 1;
		$this->_startColumn = $startColumnIndex;
		$this->seek($startColumn);

        return $this;
	}

	/**
	 * (Re)Set the end column
	 *
	 * @param string	$endColumn	The column address at which to stop iterating
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator
	 */
	public function resetEnd($endColumn = null) {
		$endColumn = ($endColumn) ? $endColumn : $this->_subject->getHighestColumn();
		$this->_endColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($endColumn) - 1;

        return $this;
	}

	/**
	 * Set the column pointer to the selected column
	 *
	 * @param string	$column	The column address to set the current pointer at
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
	public function seek($column = 'A') {
        $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column) - 1;
        if (($column < $this->_startColumn) || ($column > $this->_endColumn)) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Column $column is out of range ({$this->_startColumn} - {$this->_endColumn})");
        }
		$this->_position = $column;

        return $this;
    }

	/**
	 * Rewind the iterator to the starting column
	 */
	public function rewind() {
		$this->_position = $this->_startColumn;
	}

	/**
	 * Return the current column in this worksheet
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Worksheet\Column
	 */
	public function current() {
		return new \PhpOffice\PhpSpreadsheet\Worksheet\Column($this->_subject, \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($this->_position));
	}

	/**
	 * Return the current iterator key
	 *
	 * @return string
	 */
	public function key() {
		return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($this->_position);
	}

	/**
	 * Set the iterator to its next value
	 */
	public function next() {
		++$this->_position;
	}

	/**
	 * Set the iterator to its previous value
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
	public function prev() {
        if ($this->_position <= $this->_startColumn) {
            throw new \PhpOffice\PhpSpreadsheet\Exception(
                "Column is already at the beginning of range (" . 
                \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($this->_endColumn) . " - " . 
                \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($this->_endColumn) . ")"
            );
        }

        --$this->_position;
	}

	/**
	 * Indicate if more columns exist in the worksheet range of columns that we're iterating
	 *
	 * @return boolean
	 */
	public function valid() {
		return $this->_position <= $this->_endColumn;
	}
}

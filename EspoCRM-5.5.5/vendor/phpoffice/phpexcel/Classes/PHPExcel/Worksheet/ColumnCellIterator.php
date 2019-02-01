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
 * \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator
 *
 * Used to iterate columns in a \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator extends \PhpOffice\PhpSpreadsheet\Worksheet\CellIterator implements Iterator
{
	/**
	 * Column index
	 *
	 * @var string
	 */
	protected $_columnIndex;

    /**
	 * Start position
	 *
	 * @var int
	 */
	protected $_startRow = 1;

	/**
	 * End position
	 *
	 * @var int
	 */
	protected $_endRow = 1;

	/**
	 * Create a new row iterator
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$subject	    The worksheet to iterate over
     * @param   string              $columnIndex    The column that we want to iterate
	 * @param	integer				$startRow	    The row number at which to start iterating
	 * @param	integer				$endRow	        Optionally, the row number at which to stop iterating
	 */
	public function __construct(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $subject = null, $columnIndex, $startRow = 1, $endRow = null) {
		// Set subject
		$this->_subject = $subject;
		$this->_columnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($columnIndex) - 1;
		$this->resetEnd($endRow);
		$this->resetStart($startRow);
	}

	/**
	 * Destructor
	 */
	public function __destruct() {
		unset($this->_subject);
	}

	/**
	 * (Re)Set the start row and the current row pointer
	 *
	 * @param integer	$startRow	The row number at which to start iterating
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
	public function resetStart($startRow = 1) {
		$this->_startRow = $startRow;
        $this->adjustForExistingOnlyRange();
		$this->seek($startRow);

        return $this;
	}

	/**
	 * (Re)Set the end row
	 *
	 * @param integer	$endRow	The row number at which to stop iterating
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
	public function resetEnd($endRow = null) {
		$this->_endRow = ($endRow) ? $endRow : $this->_subject->getHighestRow();
        $this->adjustForExistingOnlyRange();

        return $this;
	}

	/**
	 * Set the row pointer to the selected row
	 *
	 * @param integer	$row	The row number to set the current pointer at
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
	public function seek($row = 1) {
        if (($row < $this->_startRow) || ($row > $this->_endRow)) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Row $row is out of range ({$this->_startRow} - {$this->_endRow})");
        } elseif ($this->_onlyExistingCells && !($this->_subject->cellExistsByColumnAndRow($this->_columnIndex, $row))) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('In "IterateOnlyExistingCells" mode and Cell does not exist');
        }
		$this->_position = $row;

        return $this;
	}

	/**
	 * Rewind the iterator to the starting row
	 */
	public function rewind() {
		$this->_position = $this->_startRow;
	}

	/**
	 * Return the current cell in this worksheet column
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Worksheet\Row
	 */
	public function current() {
		return $this->_subject->getCellByColumnAndRow($this->_columnIndex, $this->_position);
	}

	/**
	 * Return the current iterator key
	 *
	 * @return int
	 */
	public function key() {
		return $this->_position;
	}

	/**
	 * Set the iterator to its next value
	 */
	public function next() {
        do {
            ++$this->_position;
        } while (($this->_onlyExistingCells) &&
            (!$this->_subject->cellExistsByColumnAndRow($this->_columnIndex, $this->_position)) &&
            ($this->_position <= $this->_endRow));
	}

	/**
	 * Set the iterator to its previous value
	 */
	public function prev() {
        if ($this->_position <= $this->_startRow) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Row is already at the beginning of range ({$this->_startRow} - {$this->_endRow})");
        }

        do {
            --$this->_position;
        } while (($this->_onlyExistingCells) &&
            (!$this->_subject->cellExistsByColumnAndRow($this->_columnIndex, $this->_position)) &&
            ($this->_position >= $this->_startRow));
	}

	/**
	 * Indicate if more rows exist in the worksheet range of rows that we're iterating
	 *
	 * @return boolean
	 */
	public function valid() {
		return $this->_position <= $this->_endRow;
	}

	/**
	 * Validate start/end values for "IterateOnlyExistingCells" mode, and adjust if necessary
	 *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
    protected function adjustForExistingOnlyRange() {
        if ($this->_onlyExistingCells) {
            while ((!$this->_subject->cellExistsByColumnAndRow($this->_columnIndex, $this->_startRow)) &&
                ($this->_startRow <= $this->_endRow)) {
                ++$this->_startRow;
            }
            if ($this->_startRow > $this->_endRow) {
                throw new \PhpOffice\PhpSpreadsheet\Exception('No cells exist within the specified range');
            }
            while ((!$this->_subject->cellExistsByColumnAndRow($this->_columnIndex, $this->_endRow)) &&
                ($this->_endRow >= $this->_startRow)) {
                --$this->_endRow;
            }
            if ($this->_endRow < $this->_startRow) {
                throw new \PhpOffice\PhpSpreadsheet\Exception('No cells exist within the specified range');
            }
        }
    }

}

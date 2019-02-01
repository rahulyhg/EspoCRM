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
 * @version    1.8.0, 2014-03-02
 */


/**
 * \PhpOffice\PhpSpreadsheet\Worksheet\CellIterator
 *
 * Used to iterate rows in a \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
abstract class \PhpOffice\PhpSpreadsheet\Worksheet\CellIterator
{
	/**
	 * \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet to iterate
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 */
	protected $_subject;

	/**
	 * Current iterator position
	 *
	 * @var mixed
	 */
	protected $_position = null;

	/**
	 * Iterate only existing cells
	 *
	 * @var boolean
	 */
	protected $_onlyExistingCells = false;

	/**
	 * Destructor
	 */
	public function __destruct() {
		unset($this->_subject);
	}

	/**
	 * Get loop only existing cells
	 *
	 * @return boolean
	 */
    public function getIterateOnlyExistingCells() {
    	return $this->_onlyExistingCells;
    }

	/**
	 * Validate start/end values for "IterateOnlyExistingCells" mode, and adjust if necessary
	 *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
    abstract protected function adjustForExistingOnlyRange();

	/**
	 * Set the iterator to loop only existing cells
	 *
	 * @param	boolean		$value
     * @throws \PhpOffice\PhpSpreadsheet\Exception
	 */
    public function setIterateOnlyExistingCells($value = true) {
    	$this->_onlyExistingCells = (boolean) $value;

        $this->adjustForExistingOnlyRange();
    }
}

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
 * @package    \PhpOffice\PhpSpreadsheet\Writer\Xlsx
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Writer\Xlsx
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
abstract class \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
	/**
	 * Parent IWriter object
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Writer\IWriter
	 */
	private $_parentWriter;

	/**
	 * Set parent IWriter object
	 *
	 * @param \PhpOffice\PhpSpreadsheet\Writer\IWriter	$pWriter
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function setParentWriter(\PhpOffice\PhpSpreadsheet\Writer\IWriter $pWriter = null) {
		$this->_parentWriter = $pWriter;
	}

	/**
	 * Get parent IWriter object
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function getParentWriter() {
		if (!is_null($this->_parentWriter)) {
			return $this->_parentWriter;
		} else {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("No parent \PhpOffice\PhpSpreadsheet\Writer\IWriter assigned.");
		}
	}

	/**
	 * Set parent IWriter object
	 *
	 * @param \PhpOffice\PhpSpreadsheet\Writer\IWriter	$pWriter
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function __construct(\PhpOffice\PhpSpreadsheet\Writer\IWriter $pWriter = null) {
		if (!is_null($pWriter)) {
			$this->_parentWriter = $pWriter;
		}
	}

}

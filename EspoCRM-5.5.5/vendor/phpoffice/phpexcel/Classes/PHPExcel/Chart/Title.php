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
 * @category	\PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package		\PhpOffice\PhpSpreadsheet\Chart\Chart
 * @copyright	Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Chart\Title
 *
 * @category	\PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package		\PhpOffice\PhpSpreadsheet\Chart\Chart
 * @copyright	Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Chart\Title
{

	/**
	 * Title Caption
	 *
	 * @var string
	 */
	private $_caption = null;

	/**
	 * Title Layout
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Chart\Layout
	 */
	private $_layout = null;

	/**
	 * Create a new \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	public function __construct($caption = null, \PhpOffice\PhpSpreadsheet\Chart\Layout $layout = null)
	{
		$this->_caption = $caption;
		$this->_layout = $layout;
	}

	/**
	 * Get caption
	 *
	 * @return string
	 */
	public function getCaption() {
		return $this->_caption;
	}

	/**
	 * Set caption
	 *
	 * @param string $caption
     * @return \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	public function setCaption($caption = null) {
		$this->_caption = $caption;
        
        return $this;
	}

	/**
	 * Get Layout
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Layout
	 */
	public function getLayout() {
		return $this->_layout;
	}

}

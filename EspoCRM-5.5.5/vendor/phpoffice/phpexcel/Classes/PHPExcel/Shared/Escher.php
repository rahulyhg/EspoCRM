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
 * @package    \PhpOffice\PhpSpreadsheet\Shared\Escher
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * \PhpOffice\PhpSpreadsheet\Shared\Escher
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Shared\Escher
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Shared\Escher
{
	/**
	 * Drawing Group Container
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer
	 */
	private $_dggContainer;

	/**
	 * Drawing Container
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer
	 */
	private $_dgContainer;

	/**
	 * Get Drawing Group Container
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer
	 */
	public function getDggContainer()
	{
		return $this->_dggContainer;
	}

	/**
	 * Set Drawing Group Container
	 *
	 * @param \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer $dggContainer
	 */
	public function setDggContainer($dggContainer)
	{
		return $this->_dggContainer = $dggContainer;
	}

	/**
	 * Get Drawing Container
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer
	 */
	public function getDgContainer()
	{
		return $this->_dgContainer;
	}

	/**
	 * Set Drawing Container
	 *
	 * @param \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer $dgContainer
	 */
	public function setDgContainer($dgContainer)
	{
		return $this->_dgContainer = $dgContainer;
	}

}

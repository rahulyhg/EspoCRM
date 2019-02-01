<?php
/**
 * \PhpOffice\PhpSpreadsheet\Spreadsheet
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
 * @package    \PhpOffice\PhpSpreadsheet\RichText\RichText
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\RichText\Run
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\RichText\RichText
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\RichText\Run extends \PhpOffice\PhpSpreadsheet\RichText\TextElement implements \PhpOffice\PhpSpreadsheet\RichText\ITextElement
{
	/**
	 * Font
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	private $_font;

    /**
     * Create a new \PhpOffice\PhpSpreadsheet\RichText\Run instance
     *
     * @param 	string		$pText		Text
     */
    public function __construct($pText = '')
    {
    	// Initialise variables
    	$this->setText($pText);
    	$this->_font = new \PhpOffice\PhpSpreadsheet\Style\Font();
    }

	/**
	 * Get font
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function getFont() {
		return $this->_font;
	}

	/**
	 * Set font
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Font		$pFont		Font
	 * @throws 	\PhpOffice\PhpSpreadsheet\Exception
	 * @return \PhpOffice\PhpSpreadsheet\RichText\ITextElement
	 */
	public function setFont(\PhpOffice\PhpSpreadsheet\Style\Font $pFont = null) {
		$this->_font = $pFont;
		return $this;
	}

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->getText()
    		. $this->_font->getHashCode()
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

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
 * @package	\PhpOffice\PhpSpreadsheet\Style\Style
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Style\Font
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package	\PhpOffice\PhpSpreadsheet\Style\Style
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Style\Font extends \PhpOffice\PhpSpreadsheet\Style\Supervisor implements \PhpOffice\PhpSpreadsheet\IComparable
{
	/* Underline types */
	const UNDERLINE_NONE					= 'none';
	const UNDERLINE_DOUBLE					= 'double';
	const UNDERLINE_DOUBLEACCOUNTING		= 'doubleAccounting';
	const UNDERLINE_SINGLE					= 'single';
	const UNDERLINE_SINGLEACCOUNTING		= 'singleAccounting';

	/**
	 * Font Name
	 *
	 * @var string
	 */
	protected $_name			= 'Calibri';

	/**
	 * Font Size
	 *
	 * @var float
	 */
	protected $_size			= 11;

	/**
	 * Bold
	 *
	 * @var boolean
	 */
	protected $_bold			= FALSE;

	/**
	 * Italic
	 *
	 * @var boolean
	 */
	protected $_italic		= FALSE;

	/**
	 * Superscript
	 *
	 * @var boolean
	 */
	protected $_superScript	= FALSE;

	/**
	 * Subscript
	 *
	 * @var boolean
	 */
	protected $_subScript		= FALSE;

	/**
	 * Underline
	 *
	 * @var string
	 */
	protected $_underline		= self::UNDERLINE_NONE;

	/**
	 * Strikethrough
	 *
	 * @var boolean
	 */
	protected $_strikethrough	= FALSE;

	/**
	 * Foreground color
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Style\Color
	 */
	protected $_color;

	/**
	 * Create a new \PhpOffice\PhpSpreadsheet\Style\Font
	 *
	 * @param	boolean	$isSupervisor	Flag indicating if this is a supervisor or not
	 *									Leave this value at default unless you understand exactly what
	 *										its ramifications are
	 * @param	boolean	$isConditional	Flag indicating if this is a conditional style or not
	 *									Leave this value at default unless you understand exactly what
	 *										its ramifications are
	 */
	public function __construct($isSupervisor = FALSE, $isConditional = FALSE)
	{
		// Supervisor?
		parent::__construct($isSupervisor);

		// Initialise values
		if ($isConditional) {
			$this->_name			= NULL;
			$this->_size			= NULL;
			$this->_bold			= NULL;
			$this->_italic			= NULL;
			$this->_superScript		= NULL;
			$this->_subScript		= NULL;
			$this->_underline		= NULL;
			$this->_strikethrough	= NULL;
			$this->_color			= new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK, $isSupervisor, $isConditional);
		} else {
			$this->_color	= new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK, $isSupervisor);
		}
		// bind parent if we are a supervisor
		if ($isSupervisor) {
			$this->_color->bindParent($this, '_color');
		}
	}

	/**
	 * Get the shared style component for the currently active cell in currently active sheet.
	 * Only used for style supervisor
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function getSharedComponent()
	{
		return $this->_parent->getSharedComponent()->getFont();
	}

	/**
	 * Build style array from subcomponents
	 *
	 * @param array $array
	 * @return array
	 */
	public function getStyleArray($array)
	{
		return array('font' => $array);
	}

	/**
	 * Apply styles from array
	 *
	 * <code>
	 * $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('B2')->getFont()->applyFromArray(
	 *		array(
	 *			'name'		=> 'Arial',
	 *			'bold'		=> TRUE,
	 *			'italic'	=> FALSE,
	 *			'underline' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLE,
	 *			'strike'	=> FALSE,
	 *			'color'		=> array(
	 *				'rgb' => '808080'
	 *			)
	 *		)
	 * );
	 * </code>
	 *
	 * @param	array	$pStyles	Array containing style information
	 * @throws	\PhpOffice\PhpSpreadsheet\Exception
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function applyFromArray($pStyles = null) {
		if (is_array($pStyles)) {
			if ($this->_isSupervisor) {
				$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($this->getStyleArray($pStyles));
			} else {
				if (array_key_exists('name', $pStyles)) {
					$this->setName($pStyles['name']);
				}
				if (array_key_exists('bold', $pStyles)) {
					$this->setBold($pStyles['bold']);
				}
				if (array_key_exists('italic', $pStyles)) {
					$this->setItalic($pStyles['italic']);
				}
				if (array_key_exists('superScript', $pStyles)) {
					$this->setSuperScript($pStyles['superScript']);
				}
				if (array_key_exists('subScript', $pStyles)) {
					$this->setSubScript($pStyles['subScript']);
				}
				if (array_key_exists('underline', $pStyles)) {
					$this->setUnderline($pStyles['underline']);
				}
				if (array_key_exists('strike', $pStyles)) {
					$this->setStrikethrough($pStyles['strike']);
				}
				if (array_key_exists('color', $pStyles)) {
					$this->getColor()->applyFromArray($pStyles['color']);
				}
				if (array_key_exists('size', $pStyles)) {
					$this->setSize($pStyles['size']);
				}
			}
		} else {
			throw new \PhpOffice\PhpSpreadsheet\Exception("Invalid style array passed.");
		}
		return $this;
	}

	/**
	 * Get Name
	 *
	 * @return string
	 */
	public function getName() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getName();
		}
		return $this->_name;
	}

	/**
	 * Set Name
	 *
	 * @param string $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setName($pValue = 'Calibri') {
  		if ($pValue == '') {
			$pValue = 'Calibri';
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('name' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_name = $pValue;
		}
		return $this;
	}

	/**
	 * Get Size
	 *
	 * @return double
	 */
	public function getSize() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getSize();
		}
		return $this->_size;
	}

	/**
	 * Set Size
	 *
	 * @param double $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setSize($pValue = 10) {
		if ($pValue == '') {
			$pValue = 10;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('size' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_size = $pValue;
		}
		return $this;
	}

	/**
	 * Get Bold
	 *
	 * @return boolean
	 */
	public function getBold() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getBold();
		}
		return $this->_bold;
	}

	/**
	 * Set Bold
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setBold($pValue = false) {
		if ($pValue == '') {
			$pValue = false;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('bold' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_bold = $pValue;
		}
		return $this;
	}

	/**
	 * Get Italic
	 *
	 * @return boolean
	 */
	public function getItalic() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getItalic();
		}
		return $this->_italic;
	}

	/**
	 * Set Italic
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setItalic($pValue = false) {
		if ($pValue == '') {
			$pValue = false;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('italic' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_italic = $pValue;
		}
		return $this;
	}

	/**
	 * Get SuperScript
	 *
	 * @return boolean
	 */
	public function getSuperScript() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getSuperScript();
		}
		return $this->_superScript;
	}

	/**
	 * Set SuperScript
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setSuperScript($pValue = false) {
		if ($pValue == '') {
			$pValue = false;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('superScript' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_superScript = $pValue;
			$this->_subScript = !$pValue;
		}
		return $this;
	}

		/**
	 * Get SubScript
	 *
	 * @return boolean
	 */
	public function getSubScript() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getSubScript();
		}
		return $this->_subScript;
	}

	/**
	 * Set SubScript
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setSubScript($pValue = false) {
		if ($pValue == '') {
			$pValue = false;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('subScript' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_subScript = $pValue;
			$this->_superScript = !$pValue;
		}
		return $this;
	}

	/**
	 * Get Underline
	 *
	 * @return string
	 */
	public function getUnderline() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getUnderline();
		}
		return $this->_underline;
	}

	/**
	 * Set Underline
	 *
	 * @param string|boolean $pValue	\PhpOffice\PhpSpreadsheet\Style\Font underline type
	 *									If a boolean is passed, then TRUE equates to UNDERLINE_SINGLE,
	 *										false equates to UNDERLINE_NONE
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setUnderline($pValue = self::UNDERLINE_NONE) {
		if (is_bool($pValue)) {
			$pValue = ($pValue) ? self::UNDERLINE_SINGLE : self::UNDERLINE_NONE;
		} elseif ($pValue == '') {
			$pValue = self::UNDERLINE_NONE;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('underline' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_underline = $pValue;
		}
		return $this;
	}

	/**
	 * Get Strikethrough
	 *
	 * @return boolean
	 */
	public function getStrikethrough() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getStrikethrough();
		}
		return $this->_strikethrough;
	}

	/**
	 * Set Strikethrough
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setStrikethrough($pValue = false) {
		if ($pValue == '') {
			$pValue = false;
		}
		if ($this->_isSupervisor) {
			$styleArray = $this->getStyleArray(array('strike' => $pValue));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_strikethrough = $pValue;
		}
		return $this;
	}

	/**
	 * Get Color
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Style\Color
	 */
	public function getColor() {
		return $this->_color;
	}

	/**
	 * Set Color
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Color $pValue
	 * @throws	\PhpOffice\PhpSpreadsheet\Exception
	 * @return \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	public function setColor(\PhpOffice\PhpSpreadsheet\Style\Color $pValue = null) {
		// make sure parameter is a real color and not a supervisor
		$color = $pValue->getIsSupervisor() ? $pValue->getSharedComponent() : $pValue;

		if ($this->_isSupervisor) {
			$styleArray = $this->getColor()->getStyleArray(array('argb' => $color->getARGB()));
			$this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
		} else {
			$this->_color = $color;
		}
		return $this;
	}

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
		if ($this->_isSupervisor) {
			return $this->getSharedComponent()->getHashCode();
		}
		return md5(
			  $this->_name
			. $this->_size
			. ($this->_bold ? 't' : 'f')
			. ($this->_italic ? 't' : 'f')
			. ($this->_superScript ? 't' : 'f')
			. ($this->_subScript ? 't' : 'f')
			. $this->_underline
			. ($this->_strikethrough ? 't' : 'f')
			. $this->_color->getHashCode()
			. __CLASS__
		);
	}

}

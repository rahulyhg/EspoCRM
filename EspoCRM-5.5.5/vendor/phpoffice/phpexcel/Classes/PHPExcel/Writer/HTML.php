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
 * @package	\PhpOffice\PhpSpreadsheet\Writer\Html
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Writer\Html
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package	\PhpOffice\PhpSpreadsheet\Writer\Html
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Writer\Html extends \PhpOffice\PhpSpreadsheet\Writer\BaseWriter implements \PhpOffice\PhpSpreadsheet\Writer\IWriter {
	/**
	 * \PhpOffice\PhpSpreadsheet\Spreadsheet object
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
	 */
	protected $_phpExcel;

	/**
	 * Sheet index to write
	 *
	 * @var int
	 */
	private $_sheetIndex	= 0;

	/**
	 * Images root
	 *
	 * @var string
	 */
	private $_imagesRoot	= '.';

	/**
	 * embed images, or link to images
	 *
	 * @var boolean
	 */
	private $_embedImages	= FALSE;

	/**
	 * Use inline CSS?
	 *
	 * @var boolean
	 */
	private $_useInlineCss = false;

	/**
	 * Array of CSS styles
	 *
	 * @var array
	 */
	private $_cssStyles = null;

	/**
	 * Array of column widths in points
	 *
	 * @var array
	 */
	private $_columnWidths = null;

	/**
	 * Default font
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Style\Font
	 */
	private $_defaultFont;

	/**
	 * Flag whether spans have been calculated
	 *
	 * @var boolean
	 */
	private $_spansAreCalculated	= false;

	/**
	 * Excel cells that should not be written as HTML cells
	 *
	 * @var array
	 */
	private $_isSpannedCell	= array();

	/**
	 * Excel cells that are upper-left corner in a cell merge
	 *
	 * @var array
	 */
	private $_isBaseCell	= array();

	/**
	 * Excel rows that should not be written as HTML rows
	 *
	 * @var array
	 */
	private $_isSpannedRow	= array();

	/**
	 * Is the current writer creating PDF?
	 *
	 * @var boolean
	 */
	protected $_isPdf = false;

	/**
	 * Generate the Navigation block
	 *
	 * @var boolean
	 */
	private $_generateSheetNavigationBlock = true;

	/**
	 * Create a new \PhpOffice\PhpSpreadsheet\Writer\Html
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Spreadsheet	$phpExcel	\PhpOffice\PhpSpreadsheet\Spreadsheet object
	 */
	public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel) {
		$this->_phpExcel = $phpExcel;
		$this->_defaultFont = $this->_phpExcel->getDefaultStyle()->getFont();
	}

	/**
	 * Save \PhpOffice\PhpSpreadsheet\Spreadsheet to file
	 *
	 * @param	string		$pFilename
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function save($pFilename = null) {
		// garbage collect
		$this->_phpExcel->garbageCollect();

		$saveDebugLog = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($this->_phpExcel)->getDebugLog()->getWriteDebugLog();
		\PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($this->_phpExcel)->getDebugLog()->setWriteDebugLog(FALSE);
		$saveArrayReturnType = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getArrayReturnType();
		\PhpOffice\PhpSpreadsheet\Calculation\Calculation::setArrayReturnType(\PhpOffice\PhpSpreadsheet\Calculation\Calculation::RETURN_ARRAY_AS_VALUE);

		// Build CSS
		$this->buildCSS(!$this->_useInlineCss);

		// Open file
		$fileHandle = fopen($pFilename, 'wb+');
		if ($fileHandle === false) {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Could not open file $pFilename for writing.");
		}

		// Write headers
		fwrite($fileHandle, $this->generateHTMLHeader(!$this->_useInlineCss));

		// Write navigation (tabs)
		if ((!$this->_isPdf) && ($this->_generateSheetNavigationBlock)) {
			fwrite($fileHandle, $this->generateNavigation());
		}

		// Write data
		fwrite($fileHandle, $this->generateSheetData());

		// Write footer
		fwrite($fileHandle, $this->generateHTMLFooter());

		// Close file
		fclose($fileHandle);

		\PhpOffice\PhpSpreadsheet\Calculation\Calculation::setArrayReturnType($saveArrayReturnType);
		\PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($this->_phpExcel)->getDebugLog()->setWriteDebugLog($saveDebugLog);
	}

	/**
	 * Map VAlign
	 *
	 * @param	string		$vAlign		Vertical alignment
	 * @return string
	 */
	private function _mapVAlign($vAlign) {
		switch ($vAlign) {
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM:		return 'bottom';
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP:		return 'top';
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER:
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_JUSTIFY:	return 'middle';
			default: return 'baseline';
		}
	}

	/**
	 * Map HAlign
	 *
	 * @param	string		$hAlign		Horizontal alignment
	 * @return string|false
	 */
	private function _mapHAlign($hAlign) {
		switch ($hAlign) {
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL:				return false;
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT:					return 'left';
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT:				return 'right';
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER:
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER_CONTINUOUS:	return 'center';
			case \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_JUSTIFY:				return 'justify';
			default: return false;
		}
	}

	/**
	 * Map border style
	 *
	 * @param	int		$borderStyle		Sheet index
	 * @return	string
	 */
	private function _mapBorderStyle($borderStyle) {
		switch ($borderStyle) {
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE:				return 'none';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT:				return '1px dashed';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT:			return '1px dotted';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHED:				return '1px dashed';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED:				return '1px dotted';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE:				return '3px double';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR:				return '1px solid';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM:				return '2px solid';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOT:		return '2px dashed';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOTDOT:	return '2px dotted';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED:		return '2px dashed';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_SLANTDASHDOT:		return '2px dashed';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK:				return '3px solid';
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN:				return '1px solid';
			default: return '1px solid'; // map others to thin
		}
	}

	/**
	 * Get sheet index
	 *
	 * @return int
	 */
	public function getSheetIndex() {
		return $this->_sheetIndex;
	}

	/**
	 * Set sheet index
	 *
	 * @param	int		$pValue		Sheet index
	 * @return \PhpOffice\PhpSpreadsheet\Writer\Html
	 */
	public function setSheetIndex($pValue = 0) {
		$this->_sheetIndex = $pValue;
		return $this;
	}

	/**
	 * Get sheet index
	 *
	 * @return boolean
	 */
	public function getGenerateSheetNavigationBlock() {
		return $this->_generateSheetNavigationBlock;
	}

	/**
	 * Set sheet index
	 *
	 * @param	boolean		$pValue		Flag indicating whether the sheet navigation block should be generated or not
	 * @return \PhpOffice\PhpSpreadsheet\Writer\Html
	 */
	public function setGenerateSheetNavigationBlock($pValue = true) {
		$this->_generateSheetNavigationBlock = (bool) $pValue;
		return $this;
	}

	/**
	 * Write all sheets (resets sheetIndex to NULL)
	 */
	public function writeAllSheets() {
		$this->_sheetIndex = null;
		return $this;
	}

	/**
	 * Generate HTML header
	 *
	 * @param	boolean		$pIncludeStyles		Include styles?
	 * @return	string
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function generateHTMLHeader($pIncludeStyles = false) {
		// \PhpOffice\PhpSpreadsheet\Spreadsheet object known?
		if (is_null($this->_phpExcel)) {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Internal \PhpOffice\PhpSpreadsheet\Spreadsheet object not set to an instance of an object.');
		}

		// Construct HTML
		$properties = $this->_phpExcel->getProperties();
		$html = '<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">' . PHP_EOL;
		$html .= '<!-- Generated by \PhpOffice\PhpSpreadsheet\Spreadsheet - http://www.phpexcel.net -->' . PHP_EOL;
		$html .= '<html>' . PHP_EOL;
		$html .= '  <head>' . PHP_EOL;
		$html .= '	  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">' . PHP_EOL;
		if ($properties->getTitle() > '')
			$html .= '	  <title>' . htmlspecialchars($properties->getTitle()) . '</title>' . PHP_EOL;

		if ($properties->getCreator() > '')
			$html .= '	  <meta name="author" content="' . htmlspecialchars($properties->getCreator()) . '" />' . PHP_EOL;
		if ($properties->getTitle() > '')
			$html .= '	  <meta name="title" content="' . htmlspecialchars($properties->getTitle()) . '" />' . PHP_EOL;
		if ($properties->getDescription() > '')
			$html .= '	  <meta name="description" content="' . htmlspecialchars($properties->getDescription()) . '" />' . PHP_EOL;
		if ($properties->getSubject() > '')
			$html .= '	  <meta name="subject" content="' . htmlspecialchars($properties->getSubject()) . '" />' . PHP_EOL;
		if ($properties->getKeywords() > '')
			$html .= '	  <meta name="keywords" content="' . htmlspecialchars($properties->getKeywords()) . '" />' . PHP_EOL;
		if ($properties->getCategory() > '')
			$html .= '	  <meta name="category" content="' . htmlspecialchars($properties->getCategory()) . '" />' . PHP_EOL;
		if ($properties->getCompany() > '')
			$html .= '	  <meta name="company" content="' . htmlspecialchars($properties->getCompany()) . '" />' . PHP_EOL;
		if ($properties->getManager() > '')
			$html .= '	  <meta name="manager" content="' . htmlspecialchars($properties->getManager()) . '" />' . PHP_EOL;

		if ($pIncludeStyles) {
			$html .= $this->generateStyles(true);
		}

		$html .= '  </head>' . PHP_EOL;
		$html .= '' . PHP_EOL;
		$html .= '  <body>' . PHP_EOL;

		// Return
		return $html;
	}

	/**
	 * Generate sheet data
	 *
	 * @return	string
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function generateSheetData() {
		// \PhpOffice\PhpSpreadsheet\Spreadsheet object known?
		if (is_null($this->_phpExcel)) {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Internal \PhpOffice\PhpSpreadsheet\Spreadsheet object not set to an instance of an object.');
		}

		// Ensure that Spans have been calculated?
		if (!$this->_spansAreCalculated) {
			$this->_calculateSpans();
		}

		// Fetch sheets
		$sheets = array();
		if (is_null($this->_sheetIndex)) {
			$sheets = $this->_phpExcel->getAllSheets();
		} else {
			$sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
		}

		// Construct HTML
		$html = '';

		// Loop all sheets
		$sheetId = 0;
		foreach ($sheets as $sheet) {
			// Write table header
			$html .= $this->_generateTableHeader($sheet);

			// Get worksheet dimension
			$dimension = explode(':', $sheet->calculateWorksheetDimension());
			$dimension[0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($dimension[0]);
			$dimension[0][0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($dimension[0][0]) - 1;
			$dimension[1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($dimension[1]);
			$dimension[1][0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($dimension[1][0]) - 1;

			// row min,max
			$rowMin = $dimension[0][1];
			$rowMax = $dimension[1][1];

			// calculate start of <tbody>, <thead>
			$tbodyStart = $rowMin;
			$theadStart = $theadEnd   = 0; // default: no <thead>	no </thead>
			if ($sheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
				$rowsToRepeatAtTop = $sheet->getPageSetup()->getRowsToRepeatAtTop();

				// we can only support repeating rows that start at top row
				if ($rowsToRepeatAtTop[0] == 1) {
					$theadStart = $rowsToRepeatAtTop[0];
					$theadEnd   = $rowsToRepeatAtTop[1];
					$tbodyStart = $rowsToRepeatAtTop[1] + 1;
				}
			}

			// Loop through cells
			$row = $rowMin-1;
			while($row++ < $rowMax) {
				// <thead> ?
				if ($row == $theadStart) {
					$html .= '		<thead>' . PHP_EOL;
                    $cellType = 'th';
				}

				// <tbody> ?
				if ($row == $tbodyStart) {
					$html .= '		<tbody>' . PHP_EOL;
                    $cellType = 'td';
				}

				// Write row if there are HTML table cells in it
				if ( !isset($this->_isSpannedRow[$sheet->getParent()->getIndex($sheet)][$row]) ) {
					// Start a new rowData
					$rowData = array();
					// Loop through columns
					$column = $dimension[0][0] - 1;
					while($column++ < $dimension[1][0]) {
						// Cell exists?
						if ($sheet->cellExistsByColumnAndRow($column, $row)) {
							$rowData[$column] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($column) . $row;
						} else {
							$rowData[$column] = '';
						}
					}
					$html .= $this->_generateRow($sheet, $rowData, $row - 1, $cellType);
				}

				// </thead> ?
				if ($row == $theadEnd) {
					$html .= '		</thead>' . PHP_EOL;
				}
			}
			$html .= $this->_extendRowsForChartsAndImages($sheet, $row);

			// Close table body.
			$html .= '		</tbody>' . PHP_EOL;

			// Write table footer
			$html .= $this->_generateTableFooter();

			// Writing PDF?
			if ($this->_isPdf) {
				if (is_null($this->_sheetIndex) && $sheetId + 1 < $this->_phpExcel->getSheetCount()) {
					$html .= '<div style="page-break-before:always" />';
				}
			}

			// Next sheet
			++$sheetId;
		}

		// Return
		return $html;
	}

	/**
	 * Generate sheet tabs
	 *
	 * @return	string
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function generateNavigation()
	{
		// \PhpOffice\PhpSpreadsheet\Spreadsheet object known?
		if (is_null($this->_phpExcel)) {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Internal \PhpOffice\PhpSpreadsheet\Spreadsheet object not set to an instance of an object.');
		}

		// Fetch sheets
		$sheets = array();
		if (is_null($this->_sheetIndex)) {
			$sheets = $this->_phpExcel->getAllSheets();
		} else {
			$sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
		}

		// Construct HTML
		$html = '';

		// Only if there are more than 1 sheets
		if (count($sheets) > 1) {
			// Loop all sheets
			$sheetId = 0;

			$html .= '<ul class="navigation">' . PHP_EOL;

			foreach ($sheets as $sheet) {
				$html .= '  <li class="sheet' . $sheetId . '"><a href="#sheet' . $sheetId . '">' . $sheet->getTitle() . '</a></li>' . PHP_EOL;
				++$sheetId;
			}

			$html .= '</ul>' . PHP_EOL;
		}

		return $html;
	}

	private function _extendRowsForChartsAndImages(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet, $row) {
		$rowMax = $row;
		$colMax = 'A';
		if ($this->_includeCharts) {
			foreach ($pSheet->getChartCollection() as $chart) {
				if ($chart instanceof \PhpOffice\PhpSpreadsheet\Chart\Chart) {
				    $chartCoordinates = $chart->getTopLeftPosition();
				    $chartTL = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($chartCoordinates['cell']);
					$chartCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($chartTL[0]);
					if ($chartTL[1] > $rowMax) {
						$rowMax = $chartTL[1];
						if ($chartCol > \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($colMax)) {
							$colMax = $chartTL[0];
						}
					}
				}
			}
		}

		foreach ($pSheet->getDrawingCollection() as $drawing) {
			if ($drawing instanceof \PhpOffice\PhpSpreadsheet\Worksheet\Drawing) {
			    $imageTL = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($drawing->getCoordinates());
				$imageCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($imageTL[0]);
				if ($imageTL[1] > $rowMax) {
					$rowMax = $imageTL[1];
					if ($imageCol > \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($colMax)) {
						$colMax = $imageTL[0];
					}
				}
			}
		}
		$html = '';
		$colMax++;
		while ($row < $rowMax) {
			$html .= '<tr>';
			for ($col = 'A'; $col != $colMax; ++$col) {
				$html .= '<td>';
				$html .= $this->_writeImageInCell($pSheet, $col.$row);
				if ($this->_includeCharts) {
					$html .= $this->_writeChartInCell($pSheet, $col.$row);
				}
				$html .= '</td>';
			}
			++$row;
			$html .= '</tr>';
		}
		return $html;
	}


	/**
	 * Generate image tag in cell
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$pSheet			\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 * @param	string				$coordinates	Cell coordinates
	 * @return	string
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeImageInCell(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet, $coordinates) {
		// Construct HTML
		$html = '';

		// Write images
		foreach ($pSheet->getDrawingCollection() as $drawing) {
			if ($drawing instanceof \PhpOffice\PhpSpreadsheet\Worksheet\Drawing) {
				if ($drawing->getCoordinates() == $coordinates) {
					$filename = $drawing->getPath();

					// Strip off eventual '.'
					if (substr($filename, 0, 1) == '.') {
						$filename = substr($filename, 1);
					}

					// Prepend images root
					$filename = $this->getImagesRoot() . $filename;

					// Strip off eventual '.'
					if (substr($filename, 0, 1) == '.' && substr($filename, 0, 2) != './') {
						$filename = substr($filename, 1);
					}

					// Convert UTF8 data to PCDATA
					$filename = htmlspecialchars($filename);

					$html .= PHP_EOL;
					if ((!$this->_embedImages) || ($this->_isPdf)) {
						$imageData = $filename;
					} else {
						$imageDetails = getimagesize($filename);
						if ($fp = fopen($filename,"rb", 0)) {
							$picture = fread($fp,filesize($filename));
							fclose($fp);
							// base64 encode the binary data, then break it
							// into chunks according to RFC 2045 semantics
							$base64 = chunk_split(base64_encode($picture));
							$imageData = 'data:'.$imageDetails['mime'].';base64,' . $base64;
						} else {
							$imageData = $filename;
						}
					}

					$html .= '<div style="position: relative;">';
					$html .= '<img style="position: absolute; z-index: 1; left: ' . 
                        $drawing->getOffsetX() . 'px; top: ' . $drawing->getOffsetY() . 'px; width: ' . 
                        $drawing->getWidth() . 'px; height: ' . $drawing->getHeight() . 'px;" src="' . 
                        $imageData . '" border="0" />';
					$html .= '</div>';
				}
			}
		}

		// Return
		return $html;
	}

	/**
	 * Generate chart tag in cell
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$pSheet			\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 * @param	string				$coordinates	Cell coordinates
	 * @return	string
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeChartInCell(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet, $coordinates) {
		// Construct HTML
		$html = '';

		// Write charts
		foreach ($pSheet->getChartCollection() as $chart) {
			if ($chart instanceof \PhpOffice\PhpSpreadsheet\Chart\Chart) {
			    $chartCoordinates = $chart->getTopLeftPosition();
				if ($chartCoordinates['cell'] == $coordinates) {
					$chartFileName = \PhpOffice\PhpSpreadsheet\Shared\File::sys_get_temp_dir().'/'.uniqid().'.png';
					if (!$chart->render($chartFileName)) {
						return;
					}

					$html .= PHP_EOL;
					$imageDetails = getimagesize($chartFileName);
					if ($fp = fopen($chartFileName,"rb", 0)) {
						$picture = fread($fp,filesize($chartFileName));
						fclose($fp);
						// base64 encode the binary data, then break it
						// into chunks according to RFC 2045 semantics
						$base64 = chunk_split(base64_encode($picture));
						$imageData = 'data:'.$imageDetails['mime'].';base64,' . $base64;

						$html .= '<div style="position: relative;">';
						$html .= '<img style="position: absolute; z-index: 1; left: ' . $chartCoordinates['xOffset'] . 'px; top: ' . $chartCoordinates['yOffset'] . 'px; width: ' . $imageDetails[0] . 'px; height: ' . $imageDetails[1] . 'px;" src="' . $imageData . '" border="0" />' . PHP_EOL;
						$html .= '</div>';

						unlink($chartFileName);
					}
				}
			}
		}

		// Return
		return $html;
	}

	/**
	 * Generate CSS styles
	 *
	 * @param	boolean	$generateSurroundingHTML	Generate surrounding HTML tags? (&lt;style&gt; and &lt;/style&gt;)
	 * @return	string
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function generateStyles($generateSurroundingHTML = true) {
		// \PhpOffice\PhpSpreadsheet\Spreadsheet object known?
		if (is_null($this->_phpExcel)) {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Internal \PhpOffice\PhpSpreadsheet\Spreadsheet object not set to an instance of an object.');
		}

		// Build CSS
		$css = $this->buildCSS($generateSurroundingHTML);

		// Construct HTML
		$html = '';

		// Start styles
		if ($generateSurroundingHTML) {
			$html .= '	<style type="text/css">' . PHP_EOL;
			$html .= '	  html { ' . $this->_assembleCSS($css['html']) . ' }' . PHP_EOL;
		}

		// Write all other styles
		foreach ($css as $styleName => $styleDefinition) {
			if ($styleName != 'html') {
				$html .= '	  ' . $styleName . ' { ' . $this->_assembleCSS($styleDefinition) . ' }' . PHP_EOL;
			}
		}

		// End styles
		if ($generateSurroundingHTML) {
			$html .= '	</style>' . PHP_EOL;
		}

		// Return
		return $html;
	}

	/**
	 * Build CSS styles
	 *
	 * @param	boolean	$generateSurroundingHTML	Generate surrounding HTML style? (html { })
	 * @return	array
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function buildCSS($generateSurroundingHTML = true) {
		// \PhpOffice\PhpSpreadsheet\Spreadsheet object known?
		if (is_null($this->_phpExcel)) {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Internal \PhpOffice\PhpSpreadsheet\Spreadsheet object not set to an instance of an object.');
		}

		// Cached?
		if (!is_null($this->_cssStyles)) {
			return $this->_cssStyles;
		}

		// Ensure that spans have been calculated
		if (!$this->_spansAreCalculated) {
			$this->_calculateSpans();
		}

		// Construct CSS
		$css = array();

		// Start styles
		if ($generateSurroundingHTML) {
			// html { }
			$css['html']['font-family']	  = 'Calibri, Arial, Helvetica, sans-serif';
			$css['html']['font-size']		= '11pt';
			$css['html']['background-color'] = 'white';
		}


		// table { }
		$css['table']['border-collapse']  = 'collapse';
	    if (!$this->_isPdf) {
			$css['table']['page-break-after'] = 'always';
		}

		// .gridlines td { }
		$css['.gridlines td']['border'] = '1px dotted black';
		$css['.gridlines th']['border'] = '1px dotted black';

		// .b {}
		$css['.b']['text-align'] = 'center'; // BOOL

		// .e {}
		$css['.e']['text-align'] = 'center'; // ERROR

		// .f {}
		$css['.f']['text-align'] = 'right'; // FORMULA

		// .inlineStr {}
		$css['.inlineStr']['text-align'] = 'left'; // INLINE

		// .n {}
		$css['.n']['text-align'] = 'right'; // NUMERIC

		// .s {}
		$css['.s']['text-align'] = 'left'; // STRING

		// Calculate cell style hashes
		foreach ($this->_phpExcel->getCellXfCollection() as $index => $style) {
			$css['td.style' . $index] = $this->_createCSSStyle( $style );
			$css['th.style' . $index] = $this->_createCSSStyle( $style );
		}

		// Fetch sheets
		$sheets = array();
		if (is_null($this->_sheetIndex)) {
			$sheets = $this->_phpExcel->getAllSheets();
		} else {
			$sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
		}

		// Build styles per sheet
		foreach ($sheets as $sheet) {
			// Calculate hash code
			$sheetIndex = $sheet->getParent()->getIndex($sheet);

			// Build styles
			// Calculate column widths
			$sheet->calculateColumnWidths();

			// col elements, initialize
			$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestColumn()) - 1;
			$column = -1;
			while($column++ < $highestColumnIndex) {
				$this->_columnWidths[$sheetIndex][$column] = 42; // approximation
				$css['table.sheet' . $sheetIndex . ' col.col' . $column]['width'] = '42pt';
			}

			// col elements, loop through columnDimensions and set width
			foreach ($sheet->getColumnDimensions() as $columnDimension) {
				if (($width = \PhpOffice\PhpSpreadsheet\Shared\Drawing::cellDimensionToPixels($columnDimension->getWidth(), $this->_defaultFont)) >= 0) {
					$width = \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints($width);
					$column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($columnDimension->getColumnIndex()) - 1;
					$this->_columnWidths[$sheetIndex][$column] = $width;
					$css['table.sheet' . $sheetIndex . ' col.col' . $column]['width'] = $width . 'pt';

					if ($columnDimension->getVisible() === false) {
						$css['table.sheet' . $sheetIndex . ' col.col' . $column]['visibility'] = 'collapse';
						$css['table.sheet' . $sheetIndex . ' col.col' . $column]['*display'] = 'none'; // target IE6+7
					}
				}
			}

			// Default row height
			$rowDimension = $sheet->getDefaultRowDimension();

			// table.sheetN tr { }
			$css['table.sheet' . $sheetIndex . ' tr'] = array();

			if ($rowDimension->getRowHeight() == -1) {
				$pt_height = \PhpOffice\PhpSpreadsheet\Shared\Font::getDefaultRowHeightByFont($this->_phpExcel->getDefaultStyle()->getFont());
			} else {
				$pt_height = $rowDimension->getRowHeight();
			}
			$css['table.sheet' . $sheetIndex . ' tr']['height'] = $pt_height . 'pt';
			if ($rowDimension->getVisible() === false) {
				$css['table.sheet' . $sheetIndex . ' tr']['display']	= 'none';
				$css['table.sheet' . $sheetIndex . ' tr']['visibility'] = 'hidden';
			}

			// Calculate row heights
			foreach ($sheet->getRowDimensions() as $rowDimension) {
				$row = $rowDimension->getRowIndex() - 1;

				// table.sheetN tr.rowYYYYYY { }
				$css['table.sheet' . $sheetIndex . ' tr.row' . $row] = array();

				if ($rowDimension->getRowHeight() == -1) {
					$pt_height = \PhpOffice\PhpSpreadsheet\Shared\Font::getDefaultRowHeightByFont($this->_phpExcel->getDefaultStyle()->getFont());
				} else {
					$pt_height = $rowDimension->getRowHeight();
				}
				$css['table.sheet' . $sheetIndex . ' tr.row' . $row]['height'] = $pt_height . 'pt';
				if ($rowDimension->getVisible() === false) {
					$css['table.sheet' . $sheetIndex . ' tr.row' . $row]['display'] = 'none';
					$css['table.sheet' . $sheetIndex . ' tr.row' . $row]['visibility'] = 'hidden';
				}
			}
		}

		// Cache
		if (is_null($this->_cssStyles)) {
			$this->_cssStyles = $css;
		}

		// Return
		return $css;
	}

	/**
	 * Create CSS style
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Style		$pStyle			\PhpOffice\PhpSpreadsheet\Style\Style
	 * @return	array
	 */
	private function _createCSSStyle(\PhpOffice\PhpSpreadsheet\Style\Style $pStyle) {
		// Construct CSS
		$css = '';

		// Create CSS
		$css = array_merge(
			$this->_createCSSStyleAlignment($pStyle->getAlignment())
			, $this->_createCSSStyleBorders($pStyle->getBorders())
			, $this->_createCSSStyleFont($pStyle->getFont())
			, $this->_createCSSStyleFill($pStyle->getFill())
		);

		// Return
		return $css;
	}

	/**
	 * Create CSS style (\PhpOffice\PhpSpreadsheet\Style\Alignment)
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Alignment		$pStyle			\PhpOffice\PhpSpreadsheet\Style\Alignment
	 * @return	array
	 */
	private function _createCSSStyleAlignment(\PhpOffice\PhpSpreadsheet\Style\Alignment $pStyle) {
		// Construct CSS
		$css = array();

		// Create CSS
		$css['vertical-align'] = $this->_mapVAlign($pStyle->getVertical());
		if ($textAlign = $this->_mapHAlign($pStyle->getHorizontal())) {
			$css['text-align'] = $textAlign;
			if(in_array($textAlign,array('left','right')))
				$css['padding-'.$textAlign] = (string)((int)$pStyle->getIndent() * 9).'px';
		}

		// Return
		return $css;
	}

	/**
	 * Create CSS style (\PhpOffice\PhpSpreadsheet\Style\Font)
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Font		$pStyle			\PhpOffice\PhpSpreadsheet\Style\Font
	 * @return	array
	 */
	private function _createCSSStyleFont(\PhpOffice\PhpSpreadsheet\Style\Font $pStyle) {
		// Construct CSS
		$css = array();

		// Create CSS
		if ($pStyle->getBold()) {
			$css['font-weight'] = 'bold';
		}
		if ($pStyle->getUnderline() != \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE && $pStyle->getStrikethrough()) {
			$css['text-decoration'] = 'underline line-through';
		} else if ($pStyle->getUnderline() != \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE) {
			$css['text-decoration'] = 'underline';
		} else if ($pStyle->getStrikethrough()) {
			$css['text-decoration'] = 'line-through';
		}
		if ($pStyle->getItalic()) {
			$css['font-style'] = 'italic';
		}

		$css['color']		= '#' . $pStyle->getColor()->getRGB();
		$css['font-family']	= '\'' . $pStyle->getName() . '\'';
		$css['font-size']	= $pStyle->getSize() . 'pt';

		// Return
		return $css;
	}

	/**
	 * Create CSS style (\PhpOffice\PhpSpreadsheet\Style\Borders)
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Borders		$pStyle			\PhpOffice\PhpSpreadsheet\Style\Borders
	 * @return	array
	 */
	private function _createCSSStyleBorders(\PhpOffice\PhpSpreadsheet\Style\Borders $pStyle) {
		// Construct CSS
		$css = array();

		// Create CSS
		$css['border-bottom']	= $this->_createCSSStyleBorder($pStyle->getBottom());
		$css['border-top']		= $this->_createCSSStyleBorder($pStyle->getTop());
		$css['border-left']		= $this->_createCSSStyleBorder($pStyle->getLeft());
		$css['border-right']	= $this->_createCSSStyleBorder($pStyle->getRight());

		// Return
		return $css;
	}

	/**
	 * Create CSS style (\PhpOffice\PhpSpreadsheet\Style\Border)
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Border		$pStyle			\PhpOffice\PhpSpreadsheet\Style\Border
	 * @return	string
	 */
	private function _createCSSStyleBorder(\PhpOffice\PhpSpreadsheet\Style\Border $pStyle) {
		// Create CSS
//		$css = $this->_mapBorderStyle($pStyle->getBorderStyle()) . ' #' . $pStyle->getColor()->getRGB();
		//	Create CSS - add !important to non-none border styles for merged cells  
		$borderStyle = $this->_mapBorderStyle($pStyle->getBorderStyle());  
		$css = $borderStyle . ' #' . $pStyle->getColor()->getRGB() . (($borderStyle == 'none') ? '' : ' !important'); 

		// Return
		return $css;
	}

	/**
	 * Create CSS style (\PhpOffice\PhpSpreadsheet\Style\Fill)
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Style\Fill		$pStyle			\PhpOffice\PhpSpreadsheet\Style\Fill
	 * @return	array
	 */
	private function _createCSSStyleFill(\PhpOffice\PhpSpreadsheet\Style\Fill $pStyle) {
		// Construct HTML
		$css = array();

		// Create CSS
		$value = $pStyle->getFillType() == \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE ?
			'white' : '#' . $pStyle->getStartColor()->getRGB();
		$css['background-color'] = $value;

		// Return
		return $css;
	}

	/**
	 * Generate HTML footer
	 */
	public function generateHTMLFooter() {
		// Construct HTML
		$html = '';
		$html .= '  </body>' . PHP_EOL;
		$html .= '</html>' . PHP_EOL;

		// Return
		return $html;
	}

	/**
	 * Generate table header
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$pSheet		The worksheet for the table we are writing
	 * @return	string
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _generateTableHeader($pSheet) {
		$sheetIndex = $pSheet->getParent()->getIndex($pSheet);

		// Construct HTML
		$html = '';
		$html .= $this->_setMargins($pSheet);
			
		if (!$this->_useInlineCss) {
			$gridlines = $pSheet->getShowGridlines() ? ' gridlines' : '';
			$html .= '	<table border="0" cellpadding="0" cellspacing="0" id="sheet' . $sheetIndex . '" class="sheet' . $sheetIndex . $gridlines . '">' . PHP_EOL;
		} else {
			$style = isset($this->_cssStyles['table']) ?
				$this->_assembleCSS($this->_cssStyles['table']) : '';

			if ($this->_isPdf && $pSheet->getShowGridlines()) {
				$html .= '	<table border="1" cellpadding="1" id="sheet' . $sheetIndex . '" cellspacing="1" style="' . $style . '">' . PHP_EOL;
			} else {
				$html .= '	<table border="0" cellpadding="1" id="sheet' . $sheetIndex . '" cellspacing="0" style="' . $style . '">' . PHP_EOL;
			}
		}

		// Write <col> elements
		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($pSheet->getHighestColumn()) - 1;
		$i = -1;
		while($i++ < $highestColumnIndex) {
		    if (!$this->_isPdf) {
				if (!$this->_useInlineCss) {
					$html .= '		<col class="col' . $i . '">' . PHP_EOL;
				} else {
					$style = isset($this->_cssStyles['table.sheet' . $sheetIndex . ' col.col' . $i]) ?
						$this->_assembleCSS($this->_cssStyles['table.sheet' . $sheetIndex . ' col.col' . $i]) : '';
					$html .= '		<col style="' . $style . '">' . PHP_EOL;
				}
			}
		}

		// Return
		return $html;
	}

	/**
	 * Generate table footer
	 *
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _generateTableFooter() {
		// Construct HTML
		$html = '';
		$html .= '	</table>' . PHP_EOL;

		// Return
		return $html;
	}

	/**
	 * Generate row
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$pSheet			\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 * @param	array				$pValues		Array containing cells in a row
	 * @param	int					$pRow			Row number (0-based)
	 * @return	string
	 * @throws	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _generateRow(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet, $pValues = null, $pRow = 0, $cellType = 'td') {
		if (is_array($pValues)) {
			// Construct HTML
			$html = '';

			// Sheet index
			$sheetIndex = $pSheet->getParent()->getIndex($pSheet);

			// DomPDF and breaks
			if ($this->_isPdf && count($pSheet->getBreaks()) > 0) {
				$breaks = $pSheet->getBreaks();

				// check if a break is needed before this row
				if (isset($breaks['A' . $pRow])) {
					// close table: </table>
					$html .= $this->_generateTableFooter();

					// insert page break
					$html .= '<div style="page-break-before:always" />';

					// open table again: <table> + <col> etc.
					$html .= $this->_generateTableHeader($pSheet);
				}
			}

			// Write row start
			if (!$this->_useInlineCss) {
				$html .= '		  <tr class="row' . $pRow . '">' . PHP_EOL;
			} else {
				$style = isset($this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow])
					? $this->_assembleCSS($this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow]) : '';

				$html .= '		  <tr style="' . $style . '">' . PHP_EOL;
			}

			// Write cells
			$colNum = 0;
			foreach ($pValues as $cellAddress) {
                $cell = ($cellAddress > '') ? $pSheet->getCell($cellAddress) : '';
				$coordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colNum) . ($pRow + 1);
				if (!$this->_useInlineCss) {
					$cssClass = '';
					$cssClass = 'column' . $colNum;
				} else {
					$cssClass = array();
                    if ($cellType == 'th') {
                        if (isset($this->_cssStyles['table.sheet' . $sheetIndex . ' th.column' . $colNum])) {
                            $this->_cssStyles['table.sheet' . $sheetIndex . ' th.column' . $colNum];
                        }
                    } else {
                        if (isset($this->_cssStyles['table.sheet' . $sheetIndex . ' td.column' . $colNum])) {
                            $this->_cssStyles['table.sheet' . $sheetIndex . ' td.column' . $colNum];
                        }
                    }
				}
				$colSpan = 1;
				$rowSpan = 1;

				// initialize
				$cellData = '&nbsp;';

				// \PhpOffice\PhpSpreadsheet\Cell\Cell
				if ($cell instanceof \PhpOffice\PhpSpreadsheet\Cell\Cell) {
					$cellData = '';
					if (is_null($cell->getParent())) {
						$cell->attach($pSheet);
					}
					// Value
					if ($cell->getValue() instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
						// Loop through rich text elements
						$elements = $cell->getValue()->getRichTextElements();
						foreach ($elements as $element) {
							// Rich text start?
							if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
								$cellData .= '<span style="' . $this->_assembleCSS($this->_createCSSStyleFont($element->getFont())) . '">';

								if ($element->getFont()->getSuperScript()) {
									$cellData .= '<sup>';
								} else if ($element->getFont()->getSubScript()) {
									$cellData .= '<sub>';
								}
							}

							// Convert UTF8 data to PCDATA
							$cellText = $element->getText();
							$cellData .= htmlspecialchars($cellText);

							if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
								if ($element->getFont()->getSuperScript()) {
									$cellData .= '</sup>';
								} else if ($element->getFont()->getSubScript()) {
									$cellData .= '</sub>';
								}

								$cellData .= '</span>';
							}
						}
					} else {
						if ($this->_preCalculateFormulas) {
							$cellData = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString(
								$cell->getCalculatedValue(),
								$pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getNumberFormat()->getFormatCode(),
								array($this, 'formatColor')
							);
						} else {
							$cellData = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString(
								$cell->getValue(),
								$pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getNumberFormat()->getFormatCode(),
								array($this, 'formatColor')
							);
						}
						$cellData = htmlspecialchars($cellData);
						if ($pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getFont()->getSuperScript()) {
							$cellData = '<sup>'.$cellData.'</sup>';
						} elseif ($pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getFont()->getSubScript()) {
							$cellData = '<sub>'.$cellData.'</sub>';
						}
					}

					// Converts the cell content so that spaces occuring at beginning of each new line are replaced by &nbsp;
					// Example: "  Hello\n to the world" is converted to "&nbsp;&nbsp;Hello\n&nbsp;to the world"
					$cellData = preg_replace("/(?m)(?:^|\\G) /", '&nbsp;', $cellData);

					// convert newline "\n" to '<br>'
					$cellData = nl2br($cellData);

					// Extend CSS class?
					if (!$this->_useInlineCss) {
						$cssClass .= ' style' . $cell->getXfIndex();
						$cssClass .= ' ' . $cell->getDataType();
					} else {
                        if ($cellType == 'th') {
                            if (isset($this->_cssStyles['th.style' . $cell->getXfIndex()])) {
                                $cssClass = array_merge($cssClass, $this->_cssStyles['th.style' . $cell->getXfIndex()]);
                            }
                        } else {
                            if (isset($this->_cssStyles['td.style' . $cell->getXfIndex()])) {
                                $cssClass = array_merge($cssClass, $this->_cssStyles['td.style' . $cell->getXfIndex()]);
                            }
                        }

						// General horizontal alignment: Actual horizontal alignment depends on dataType
						$sharedStyle = $pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() );
						if ($sharedStyle->getAlignment()->getHorizontal() == \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL
							&& isset($this->_cssStyles['.' . $cell->getDataType()]['text-align']))
						{
							$cssClass['text-align'] = $this->_cssStyles['.' . $cell->getDataType()]['text-align'];
						}
					}
				}

				// Hyperlink?
				if ($pSheet->hyperlinkExists($coordinate) && !$pSheet->getHyperlink($coordinate)->isInternal()) {
					$cellData = '<a href="' . htmlspecialchars($pSheet->getHyperlink($coordinate)->getUrl()) . '" title="' . htmlspecialchars($pSheet->getHyperlink($coordinate)->getTooltip()) . '">' . $cellData . '</a>';
				}

				// Should the cell be written or is it swallowed by a rowspan or colspan?
				$writeCell = ! ( isset($this->_isSpannedCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum])
							&& $this->_isSpannedCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum] );

				// Colspan and Rowspan
				$colspan = 1;
				$rowspan = 1;
				if (isset($this->_isBaseCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum])) {
					$spans = $this->_isBaseCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum];
					$rowSpan = $spans['rowspan'];
					$colSpan = $spans['colspan'];

					//	Also apply style from last cell in merge to fix borders -
					//		relies on !important for non-none border declarations in _createCSSStyleBorder
					$endCellCoord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colNum + $colSpan - 1) . ($pRow + $rowSpan);
					if (!$this->_useInlineCss) {
						$cssClass .= ' style' . $pSheet->getCell($endCellCoord)->getXfIndex();
					}
				}

				// Write
				if ($writeCell) {
					// Column start
					$html .= '			<' . $cellType;
						if (!$this->_useInlineCss) {
							$html .= ' class="' . $cssClass . '"';
						} else {
							//** Necessary redundant code for the sake of \PhpOffice\PhpSpreadsheet\Writer\Pdf **
							// We must explicitly write the width of the <td> element because TCPDF
							// does not recognize e.g. <col style="width:42pt">
							$width = 0;
							$i = $colNum - 1;
							$e = $colNum + $colSpan - 1;
							while($i++ < $e) {
								if (isset($this->_columnWidths[$sheetIndex][$i])) {
									$width += $this->_columnWidths[$sheetIndex][$i];
								}
							}
							$cssClass['width'] = $width . 'pt';

							// We must also explicitly write the height of the <td> element because TCPDF
							// does not recognize e.g. <tr style="height:50pt">
							if (isset($this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow]['height'])) {
								$height = $this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow]['height'];
								$cssClass['height'] = $height;
							}
							//** end of redundant code **

							$html .= ' style="' . $this->_assembleCSS($cssClass) . '"';
						}
						if ($colSpan > 1) {
							$html .= ' colspan="' . $colSpan . '"';
						}
						if ($rowSpan > 1) {
							$html .= ' rowspan="' . $rowSpan . '"';
						}
					$html .= '>';

					// Image?
					$html .= $this->_writeImageInCell($pSheet, $coordinate);

					// Chart?
					if ($this->_includeCharts) {
						$html .= $this->_writeChartInCell($pSheet, $coordinate);
					}

					// Cell data
					$html .= $cellData;

					// Column end
					$html .= '</'.$cellType.'>' . PHP_EOL;
				}

				// Next column
				++$colNum;
			}

			// Write row end
			$html .= '		  </tr>' . PHP_EOL;

			// Return
			return $html;
		} else {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
		}
	}

	/**
	 * Takes array where of CSS properties / values and converts to CSS string
	 *
	 * @param array
	 * @return string
	 */
	private function _assembleCSS($pValue = array())
	{
		$pairs = array();
		foreach ($pValue as $property => $value) {
			$pairs[] = $property . ':' . $value;
		}
		$string = implode('; ', $pairs);

		return $string;
	}

	/**
	 * Get images root
	 *
	 * @return string
	 */
	public function getImagesRoot() {
		return $this->_imagesRoot;
	}

	/**
	 * Set images root
	 *
	 * @param string $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Writer\Html
	 */
	public function setImagesRoot($pValue = '.') {
		$this->_imagesRoot = $pValue;
		return $this;
	}

	/**
	 * Get embed images
	 *
	 * @return boolean
	 */
	public function getEmbedImages() {
		return $this->_embedImages;
	}

	/**
	 * Set embed images
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Writer\Html
	 */
	public function setEmbedImages($pValue = '.') {
		$this->_embedImages = $pValue;
		return $this;
	}

	/**
	 * Get use inline CSS?
	 *
	 * @return boolean
	 */
	public function getUseInlineCss() {
		return $this->_useInlineCss;
	}

	/**
	 * Set use inline CSS?
	 *
	 * @param boolean $pValue
	 * @return \PhpOffice\PhpSpreadsheet\Writer\Html
	 */
	public function setUseInlineCss($pValue = false) {
		$this->_useInlineCss = $pValue;
		return $this;
	}

	/**
	 * Add color to formatted string as inline style
	 *
	 * @param string $pValue Plain formatted value without color
	 * @param string $pFormat Format code
	 * @return string
	 */
	public function formatColor($pValue, $pFormat)
	{
		// Color information, e.g. [Red] is always at the beginning
		$color = null; // initialize
		$matches = array();

		$color_regex = '/^\\[[a-zA-Z]+\\]/';
		if (preg_match($color_regex, $pFormat, $matches)) {
			$color = str_replace('[', '', $matches[0]);
			$color = str_replace(']', '', $color);
			$color = strtolower($color);
		}

		// convert to PCDATA
		$value = htmlspecialchars($pValue);

		// color span tag
		if ($color !== null) {
			$value = '<span style="color:' . $color . '">' . $value . '</span>';
		}

		return $value;
	}

	/**
	 * Calculate information about HTML colspan and rowspan which is not always the same as Excel's
	 */
	private function _calculateSpans()
	{
		// Identify all cells that should be omitted in HTML due to cell merge.
		// In HTML only the upper-left cell should be written and it should have
		//   appropriate rowspan / colspan attribute
		$sheetIndexes = $this->_sheetIndex !== null ?
			array($this->_sheetIndex) : range(0, $this->_phpExcel->getSheetCount() - 1);

		foreach ($sheetIndexes as $sheetIndex) {
			$sheet = $this->_phpExcel->getSheet($sheetIndex);

			$candidateSpannedRow  = array();

			// loop through all Excel merged cells
			foreach ($sheet->getMergeCells() as $cells) {
				list($cells, ) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($cells);
				$first = $cells[0];
				$last  = $cells[1];

				list($fc, $fr) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($first);
				$fc = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($fc) - 1;

				list($lc, $lr) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($last);
				$lc = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($lc) - 1;

				// loop through the individual cells in the individual merge
				$r = $fr - 1;
				while($r++ < $lr) {
					// also, flag this row as a HTML row that is candidate to be omitted
					$candidateSpannedRow[$r] = $r;

					$c = $fc - 1;
					while($c++ < $lc) {
						if ( !($c == $fc && $r == $fr) ) {
							// not the upper-left cell (should not be written in HTML)
							$this->_isSpannedCell[$sheetIndex][$r][$c] = array(
								'baseCell' => array($fr, $fc),
							);
						} else {
							// upper-left is the base cell that should hold the colspan/rowspan attribute
							$this->_isBaseCell[$sheetIndex][$r][$c] = array(
								'xlrowspan' => $lr - $fr + 1, // Excel rowspan
								'rowspan'   => $lr - $fr + 1, // HTML rowspan, value may change
								'xlcolspan' => $lc - $fc + 1, // Excel colspan
								'colspan'   => $lc - $fc + 1, // HTML colspan, value may change
							);
						}
					}
				}
			}

			// Identify which rows should be omitted in HTML. These are the rows where all the cells
			//   participate in a merge and the where base cells are somewhere above.
			$countColumns = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestColumn());
			foreach ($candidateSpannedRow as $rowIndex) {
				if (isset($this->_isSpannedCell[$sheetIndex][$rowIndex])) {
					if (count($this->_isSpannedCell[$sheetIndex][$rowIndex]) == $countColumns) {
						$this->_isSpannedRow[$sheetIndex][$rowIndex] = $rowIndex;
					};
				}
			}

			// For each of the omitted rows we found above, the affected rowspans should be subtracted by 1
			if ( isset($this->_isSpannedRow[$sheetIndex]) ) {
				foreach ($this->_isSpannedRow[$sheetIndex] as $rowIndex) {
					$adjustedBaseCells = array();
					$c = -1;
					$e = $countColumns - 1;
					while($c++ < $e) {
						$baseCell = $this->_isSpannedCell[$sheetIndex][$rowIndex][$c]['baseCell'];

						if ( !in_array($baseCell, $adjustedBaseCells) ) {
							// subtract rowspan by 1
							--$this->_isBaseCell[$sheetIndex][ $baseCell[0] ][ $baseCell[1] ]['rowspan'];
							$adjustedBaseCells[] = $baseCell;
						}
					}
				}
			}

			// TODO: Same for columns
		}

		// We have calculated the spans
		$this->_spansAreCalculated = true;
	}

	private function _setMargins(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet) {
		$htmlPage = '@page { ';
		$htmlBody = 'body { ';

		$left = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($pSheet->getPageMargins()->getLeft()) . 'in; ';
		$htmlPage .= 'left-margin: ' . $left;
		$htmlBody .= 'left-margin: ' . $left;
		$right = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($pSheet->getPageMargins()->getRight()) . 'in; ';
		$htmlPage .= 'right-margin: ' . $right;
		$htmlBody .= 'right-margin: ' . $right;
		$top = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($pSheet->getPageMargins()->getTop()) . 'in; ';
		$htmlPage .= 'top-margin: ' . $top;
		$htmlBody .= 'top-margin: ' . $top;
		$bottom = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($pSheet->getPageMargins()->getBottom()) . 'in; ';
		$htmlPage .= 'bottom-margin: ' . $bottom;
		$htmlBody .= 'bottom-margin: ' . $bottom;

		$htmlPage .= "}\n";
		$htmlBody .= "}\n";

		return "<style>\n" . $htmlPage . $htmlBody . "</style>\n";
	}
	
}

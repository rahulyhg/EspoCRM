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
 * \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Workbook
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Writer\Xlsx
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Workbook extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
	/**
	 * Write workbook to XML format
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet	$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @param	boolean		$recalcRequired	Indicate whether formulas should be recalculated before writing
	 * @return 	string 		XML Output
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeWorkbook(\PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null, $recalcRequired = FALSE)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// workbook
		$objWriter->startElement('workbook');
		$objWriter->writeAttribute('xml:space', 'preserve');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

			// fileVersion
			$this->_writeFileVersion($objWriter);

			// workbookPr
			$this->_writeWorkbookPr($objWriter);

			// workbookProtection
			$this->_writeWorkbookProtection($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet);

			// bookViews
			if ($this->getParentWriter()->getOffice2003Compatibility() === false) {
				$this->_writeBookViews($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet);
			}

			// sheets
			$this->_writeSheets($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet);

			// definedNames
			$this->_writeDefinedNames($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet);

			// calcPr
			$this->_writeCalcPr($objWriter,$recalcRequired);

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write file version
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter 		XML Writer
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeFileVersion(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null)
	{
		$objWriter->startElement('fileVersion');
		$objWriter->writeAttribute('appName', 'xl');
		$objWriter->writeAttribute('lastEdited', '4');
		$objWriter->writeAttribute('lowestEdited', '4');
		$objWriter->writeAttribute('rupBuild', '4505');
		$objWriter->endElement();
	}

	/**
	 * Write WorkbookPr
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter 		XML Writer
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeWorkbookPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null)
	{
		$objWriter->startElement('workbookPr');

		if (\PhpOffice\PhpSpreadsheet\Shared\Date::getExcelCalendar() == \PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_MAC_1904) {
			$objWriter->writeAttribute('date1904', '1');
		}

		$objWriter->writeAttribute('codeName', 'ThisWorkbook');

		$objWriter->endElement();
	}

	/**
	 * Write BookViews
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter 	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet					$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeBookViews(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
	{
		// bookViews
		$objWriter->startElement('bookViews');

			// workbookView
			$objWriter->startElement('workbookView');

			$objWriter->writeAttribute('activeTab', $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheetIndex());
			$objWriter->writeAttribute('autoFilterDateGrouping', '1');
			$objWriter->writeAttribute('firstSheet', '0');
			$objWriter->writeAttribute('minimized', '0');
			$objWriter->writeAttribute('showHorizontalScroll', '1');
			$objWriter->writeAttribute('showSheetTabs', '1');
			$objWriter->writeAttribute('showVerticalScroll', '1');
			$objWriter->writeAttribute('tabRatio', '600');
			$objWriter->writeAttribute('visibility', 'visible');

			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write WorkbookProtection
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter 	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet					$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeWorkbookProtection(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
	{
		if ($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->isSecurityEnabled()) {
			$objWriter->startElement('workbookProtection');
			$objWriter->writeAttribute('lockRevision',		($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getLockRevision() ? 'true' : 'false'));
			$objWriter->writeAttribute('lockStructure', 	($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getLockStructure() ? 'true' : 'false'));
			$objWriter->writeAttribute('lockWindows', 		($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getLockWindows() ? 'true' : 'false'));

			if ($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getRevisionsPassword() != '') {
				$objWriter->writeAttribute('revisionsPassword',	$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getRevisionsPassword());
			}

			if ($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getWorkbookPassword() != '') {
				$objWriter->writeAttribute('workbookPassword',	$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->getWorkbookPassword());
			}

			$objWriter->endElement();
		}
	}

	/**
	 * Write calcPr
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter		XML Writer
	 * @param	boolean						$recalcRequired	Indicate whether formulas should be recalculated before writing
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeCalcPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, $recalcRequired = TRUE)
	{
		$objWriter->startElement('calcPr');

		//	Set the calcid to a higher value than Excel itself will use, otherwise Excel will always recalc
        //  If MS Excel does do a recalc, then users opening a file in MS Excel will be prompted to save on exit
        //     because the file has changed
		$objWriter->writeAttribute('calcId', 			'999999');
		$objWriter->writeAttribute('calcMode', 			'auto');
		//	fullCalcOnLoad isn't needed if we've recalculating for the save
		$objWriter->writeAttribute('calcCompleted', 	($recalcRequired) ? 1 : 0);
		$objWriter->writeAttribute('fullCalcOnLoad', 	($recalcRequired) ? 0 : 1);

		$objWriter->endElement();
	}

	/**
	 * Write sheets
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter 	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet					$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeSheets(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
	{
		// Write sheets
		$objWriter->startElement('sheets');
		$sheetCount = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheetCount();
		for ($i = 0; $i < $sheetCount; ++$i) {
			// sheet
			$this->_writeSheet(
				$objWriter,
				$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheet($i)->getTitle(),
				($i + 1),
				($i + 1 + 3),
				$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheet($i)->getSheetState()
			);
		}

		$objWriter->endElement();
	}

	/**
	 * Write sheet
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter 	$objWriter 		XML Writer
	 * @param 	string 						$pSheetname 		Sheet name
	 * @param 	int							$pSheetId	 		Sheet id
	 * @param 	int							$pRelId				Relationship ID
	 * @param   string                      $sheetState         Sheet state (visible, hidden, veryHidden)
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeSheet(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, $pSheetname = '', $pSheetId = 1, $pRelId = 1, $sheetState = 'visible')
	{
		if ($pSheetname != '') {
			// Write sheet
			$objWriter->startElement('sheet');
			$objWriter->writeAttribute('name', 		$pSheetname);
			$objWriter->writeAttribute('sheetId', 	$pSheetId);
			if ($sheetState != 'visible' && $sheetState != '') {
				$objWriter->writeAttribute('state', $sheetState);
			}
			$objWriter->writeAttribute('r:id', 		'rId' . $pRelId);
			$objWriter->endElement();
		} else {
			throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
		}
	}

	/**
	 * Write Defined Names
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet					$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeDefinedNames(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
	{
		// Write defined names
		$objWriter->startElement('definedNames');

		// Named ranges
		if (count($p\PhpOffice\PhpSpreadsheet\Spreadsheet->getNamedRanges()) > 0) {
			// Named ranges
			$this->_writeNamedRanges($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet);
		}

		// Other defined names
		$sheetCount = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheetCount();
		for ($i = 0; $i < $sheetCount; ++$i) {
			// definedName for autoFilter
			$this->_writeDefinedNameForAutofilter($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheet($i), $i);

			// definedName for Print_Titles
			$this->_writeDefinedNameForPrintTitles($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheet($i), $i);

			// definedName for Print_Area
			$this->_writeDefinedNameForPrintArea($objWriter, $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheet($i), $i);
		}

		$objWriter->endElement();
	}

	/**
	 * Write named ranges
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet					$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeNamedRanges(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet)
	{
		// Loop named ranges
		$namedRanges = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getNamedRanges();
		foreach ($namedRanges as $namedRange) {
			$this->_writeDefinedNameForNamedRange($objWriter, $namedRange);
		}
	}

	/**
	 * Write Defined Name for named range
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\NamedRange			$pNamedRange
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeDefinedNameForNamedRange(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\NamedRange $pNamedRange)
	{
		// definedName for named range
		$objWriter->startElement('definedName');
		$objWriter->writeAttribute('name',			$pNamedRange->getName());
		if ($pNamedRange->getLocalOnly()) {
			$objWriter->writeAttribute('localSheetId',	$pNamedRange->getScope()->getParent()->getIndex($pNamedRange->getScope()));
		}

		// Create absolute coordinate and write as raw text
		$range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($pNamedRange->getRange());
		for ($i = 0; $i < count($range); $i++) {
			$range[$i][0] = '\'' . str_replace("'", "''", $pNamedRange->getWorksheet()->getTitle()) . '\'!' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($range[$i][0]);
			if (isset($range[$i][1])) {
				$range[$i][1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($range[$i][1]);
			}
		}
		$range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::buildRange($range);

		$objWriter->writeRawData($range);

		$objWriter->endElement();
	}

	/**
	 * Write Defined Name for autoFilter
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet			$pSheet
	 * @param 	int							$pSheetId
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeDefinedNameForAutofilter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet = null, $pSheetId = 0)
	{
		// definedName for autoFilter
		$autoFilterRange = $pSheet->getAutoFilter()->getRange();
		if (!empty($autoFilterRange)) {
			$objWriter->startElement('definedName');
			$objWriter->writeAttribute('name',			'_xlnm._FilterDatabase');
			$objWriter->writeAttribute('localSheetId',	$pSheetId);
			$objWriter->writeAttribute('hidden',		'1');

			// Create absolute coordinate and write as raw text
			$range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($autoFilterRange);
			$range = $range[0];
			//	Strip any worksheet ref so we can make the cell ref absolute
			if (strpos($range[0],'!') !== false) {
				list($ws,$range[0]) = explode('!',$range[0]);
			}

			$range[0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteCoordinate($range[0]);
			$range[1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteCoordinate($range[1]);
			$range = implode(':', $range);

			$objWriter->writeRawData('\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!' . $range);

			$objWriter->endElement();
		}
	}

	/**
	 * Write Defined Name for PrintTitles
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet			$pSheet
	 * @param 	int							$pSheetId
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeDefinedNameForPrintTitles(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet = null, $pSheetId = 0)
	{
		// definedName for PrintTitles
		if ($pSheet->getPageSetup()->isColumnsToRepeatAtLeftSet() || $pSheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
			$objWriter->startElement('definedName');
			$objWriter->writeAttribute('name',			'_xlnm.Print_Titles');
			$objWriter->writeAttribute('localSheetId',	$pSheetId);

			// Setting string
			$settingString = '';

			// Columns to repeat
			if ($pSheet->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
				$repeat = $pSheet->getPageSetup()->getColumnsToRepeatAtLeft();

				$settingString .= '\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!$' . $repeat[0] . ':$' . $repeat[1];
			}

			// Rows to repeat
			if ($pSheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
				if ($pSheet->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
					$settingString .= ',';
				}

				$repeat = $pSheet->getPageSetup()->getRowsToRepeatAtTop();

				$settingString .= '\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!$' . $repeat[0] . ':$' . $repeat[1];
			}

			$objWriter->writeRawData($settingString);

			$objWriter->endElement();
		}
	}

	/**
	 * Write Defined Name for PrintTitles
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Shared\XMLWriter	$objWriter 		XML Writer
	 * @param 	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet			$pSheet
	 * @param 	int							$pSheetId
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function _writeDefinedNameForPrintArea(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pSheet = null, $pSheetId = 0)
	{
		// definedName for PrintArea
		if ($pSheet->getPageSetup()->isPrintAreaSet()) {
			$objWriter->startElement('definedName');
			$objWriter->writeAttribute('name',			'_xlnm.Print_Area');
			$objWriter->writeAttribute('localSheetId',	$pSheetId);

			// Setting string
			$settingString = '';

			// Print area
			$printArea = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($pSheet->getPageSetup()->getPrintArea());

			$chunks = array();
			foreach ($printArea as $printAreaRect) {
				$printAreaRect[0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($printAreaRect[0]);
				$printAreaRect[1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($printAreaRect[1]);
				$chunks[] = '\'' . str_replace("'", "''", $pSheet->getTitle()) . '\'!' . implode(':', $printAreaRect);
			}

			$objWriter->writeRawData(implode(',', $chunks));

			$objWriter->endElement();
		}
	}
}

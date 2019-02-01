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
 * \PhpOffice\PhpSpreadsheet\Writer\Xlsx\DocProps
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Writer\Xlsx
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Writer\Xlsx\DocProps extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
/**
	 * Write docProps/app.xml to XML format
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet	$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @return 	string 		XML Output
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeDocPropsApp(\PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
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

		// Properties
		$objWriter->startElement('Properties');
			$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
			$objWriter->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

			// Application
			$objWriter->writeElement('Application', 	'Microsoft Excel');

			// DocSecurity
			$objWriter->writeElement('DocSecurity', 	'0');

			// ScaleCrop
			$objWriter->writeElement('ScaleCrop', 		'false');

			// HeadingPairs
			$objWriter->startElement('HeadingPairs');

				// Vector
				$objWriter->startElement('vt:vector');
					$objWriter->writeAttribute('size', 		'2');
					$objWriter->writeAttribute('baseType', 	'variant');

					// Variant
					$objWriter->startElement('vt:variant');
						$objWriter->writeElement('vt:lpstr', 	'Worksheets');
					$objWriter->endElement();

					// Variant
					$objWriter->startElement('vt:variant');
						$objWriter->writeElement('vt:i4', 		$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheetCount());
					$objWriter->endElement();

				$objWriter->endElement();

			$objWriter->endElement();

			// TitlesOfParts
			$objWriter->startElement('TitlesOfParts');

				// Vector
				$objWriter->startElement('vt:vector');
					$objWriter->writeAttribute('size', 		$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheetCount());
					$objWriter->writeAttribute('baseType',	'lpstr');

					$sheetCount = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheetCount();
					for ($i = 0; $i < $sheetCount; ++$i) {
						$objWriter->writeElement('vt:lpstr', $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getSheet($i)->getTitle());
					}

				$objWriter->endElement();

			$objWriter->endElement();

			// Company
			$objWriter->writeElement('Company', 			$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCompany());

			// Company
			$objWriter->writeElement('Manager', 			$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getManager());

			// LinksUpToDate
			$objWriter->writeElement('LinksUpToDate', 		'false');

			// SharedDoc
			$objWriter->writeElement('SharedDoc', 			'false');

			// HyperlinksChanged
			$objWriter->writeElement('HyperlinksChanged', 	'false');

			// AppVersion
			$objWriter->writeElement('AppVersion', 			'12.0000');

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write docProps/core.xml to XML format
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet	$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @return 	string 		XML Output
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeDocPropsCore(\PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
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

		// cp:coreProperties
		$objWriter->startElement('cp:coreProperties');
			$objWriter->writeAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
			$objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
			$objWriter->writeAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
			$objWriter->writeAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
			$objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

			// dc:creator
			$objWriter->writeElement('dc:creator',			$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCreator());

			// cp:lastModifiedBy
			$objWriter->writeElement('cp:lastModifiedBy', 	$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getLastModifiedBy());

			// dcterms:created
			$objWriter->startElement('dcterms:created');
				$objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
				$objWriter->writeRawData(date(DATE_W3C, 	$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCreated()));
			$objWriter->endElement();

			// dcterms:modified
			$objWriter->startElement('dcterms:modified');
				$objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
				$objWriter->writeRawData(date(DATE_W3C, 	$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getModified()));
			$objWriter->endElement();

			// dc:title
			$objWriter->writeElement('dc:title', 			$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getTitle());

			// dc:description
			$objWriter->writeElement('dc:description', 		$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getDescription());

			// dc:subject
			$objWriter->writeElement('dc:subject', 			$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getSubject());

			// cp:keywords
			$objWriter->writeElement('cp:keywords', 		$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getKeywords());

			// cp:category
			$objWriter->writeElement('cp:category', 		$p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCategory());

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write docProps/custom.xml to XML format
	 *
	 * @param 	\PhpOffice\PhpSpreadsheet\Spreadsheet	$p\PhpOffice\PhpSpreadsheet\Spreadsheet
	 * @return 	string 		XML Output
	 * @throws 	\PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function writeDocPropsCustom(\PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
	{
		$customPropertyList = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCustomProperties();
		if (empty($customPropertyList)) {
			return;
		}

		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// cp:coreProperties
		$objWriter->startElement('Properties');
			$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties');
			$objWriter->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');


			foreach($customPropertyList as $key => $customProperty) {
				$propertyValue = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCustomPropertyValue($customProperty);
				$propertyType = $p\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCustomPropertyType($customProperty);

				$objWriter->startElement('property');
					$objWriter->writeAttribute('fmtid', 	'{D5CDD505-2E9C-101B-9397-08002B2CF9AE}');
					$objWriter->writeAttribute('pid', 		$key+2);
					$objWriter->writeAttribute('name', 		$customProperty);

					switch($propertyType) {
						case 'i' :
							$objWriter->writeElement('vt:i4', 		$propertyValue);
							break;
						case 'f' :
							$objWriter->writeElement('vt:r8', 		$propertyValue);
							break;
						case 'b' :
							$objWriter->writeElement('vt:bool', 	($propertyValue) ? 'true' : 'false');
							break;
						case 'd' :
							$objWriter->startElement('vt:filetime');
								$objWriter->writeRawData(date(DATE_W3C, $propertyValue));
							$objWriter->endElement();
							break;
						default :
							$objWriter->writeElement('vt:lpwstr', 	$propertyValue);
							break;
					}

				$objWriter->endElement();
			}


		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

}

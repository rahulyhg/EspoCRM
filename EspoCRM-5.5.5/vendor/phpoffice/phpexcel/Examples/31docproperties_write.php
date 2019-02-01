<?php
/**
 * \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 * Copyright (C) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet
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
 * @package    \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/** Include \PhpOffice\PhpSpreadsheet\Spreadsheet */
require_once dirname(__FILE__) . '/../Classes/\PhpOffice\PhpSpreadsheet\Spreadsheet.php';


$inputFileType = 'Excel2007';
$inputFileName = 'templates/31docproperties.xlsx';


echo date('H:i:s') , " Load Tests from $inputFileType file" , EOL;
$callStartTime = microtime(true);

$obj\PhpOffice\PhpSpreadsheet\SpreadsheetReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet = $obj\PhpOffice\PhpSpreadsheet\SpreadsheetReader->load($inputFileName);

$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;
echo 'Call time to read Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


echo date('H:i:s') , " Adjust properties" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->setTitle("Office 2007 XLSX Test Document")
							 ->setSubject("Office 2007 XLSX Test Document")
							 ->setDescription("Test XLSX document, generated using \PhpOffice\PhpSpreadsheet\Spreadsheet")
							 ->setKeywords("office 2007 openxml php");


// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($obj\PhpOffice\PhpSpreadsheet\Spreadsheet, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB" , EOL;


echo EOL;
// Reread File
echo date('H:i:s') , " Reread Excel2007 file" , EOL;
$obj\PhpOffice\PhpSpreadsheet\SpreadsheetRead = \PhpOffice\PhpSpreadsheet\IOFactory::load(str_replace('.php', '.xlsx', __FILE__));

// Set properties
echo date('H:i:s') , " Get properties" , EOL;

echo 'Core Properties:' , EOL;
echo '    Created by - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCreator() , EOL;
echo '    Created on - ' , date('d-M-Y',$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCreated()) , ' at ' ,
                       date('H:i:s',$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCreated()) , EOL;
echo '    Last Modified by - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getLastModifiedBy() , EOL;
echo '    Last Modified on - ' , date('d-M-Y',$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getModified()) , ' at ' ,
                             date('H:i:s',$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getModified()) , EOL;
echo '    Title - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getTitle() , EOL;
echo '    Subject - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getSubject() , EOL;
echo '    Description - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getDescription() , EOL;
echo '    Keywords: - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getKeywords() , EOL;


echo 'Extended (Application) Properties:' , EOL;
echo '    Category - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCategory() , EOL;
echo '    Company - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCompany() , EOL;
echo '    Manager - ' , $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getManager() , EOL;


echo 'Custom Properties:' , EOL;
$customProperties = $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCustomProperties();
foreach($customProperties as $customProperty) {
	$propertyValue = $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCustomPropertyValue($customProperty);
	$propertyType = $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getCustomPropertyType($customProperty);
	echo '    ' , $customProperty , ' - (' , $propertyType , ') - ';
	if ($propertyType == \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_DATE) {
		echo date('d-M-Y H:i:s',$propertyValue) , EOL;
	} elseif ($propertyType == \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_BOOLEAN) {
		echo (($propertyValue) ? 'TRUE' : 'FALSE') , EOL;
	} else {
		echo $propertyValue , EOL;
	}
}

// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) . " MB" , EOL;

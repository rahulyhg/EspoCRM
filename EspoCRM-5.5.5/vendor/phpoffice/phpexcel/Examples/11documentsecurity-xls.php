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


// Create new \PhpOffice\PhpSpreadsheet\Spreadsheet object
echo date('H:i:s') , " Create new \PhpOffice\PhpSpreadsheet\Spreadsheet object" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();

// Set document properties
echo date('H:i:s') , " Set document properties" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("Office 2007 XLSX Test Document")
							 ->setSubject("Office 2007 XLSX Test Document")
							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
							 ->setKeywords("office 2007 openxml php")
							 ->setCategory("Test result file");


// Add some data
echo date('H:i:s') , " Add some data" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A1', 'Hello');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B2', 'world!');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C1', 'Hello');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('D2', 'world!');

// Rename worksheet
echo date('H:i:s') , " Rename worksheet" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setTitle('Simple');


// Set document security
echo date('H:i:s') , " Set document security" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->setLockWindows(true);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->setLockStructure(true);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getSecurity()->setWorkbookPassword("\PhpOffice\PhpSpreadsheet\Spreadsheet");


// Set sheet security
echo date('H:i:s') , " Set sheet security" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getProtection()->setPassword('\PhpOffice\PhpSpreadsheet\Spreadsheet');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getProtection()->setSheet(true); // This should be enabled in order to enable any of the following!
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getProtection()->setSort(true);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getProtection()->setInsertRows(true);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getProtection()->setFormatCells(true);


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0);


// Save Excel 95 file
echo date('H:i:s') , " Write to Excel5 format" , EOL;
$callStartTime = microtime(true);

$objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($obj\PhpOffice\PhpSpreadsheet\Spreadsheet, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

echo date('H:i:s') , " File written to " , str_replace('.php', '.xls', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;

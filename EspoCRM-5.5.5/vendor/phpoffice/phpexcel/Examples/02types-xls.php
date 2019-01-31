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
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

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

// Set default font
echo date('H:i:s') , " Set default font" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getDefaultStyle()->getFont()->setName('Arial')
                                          ->setSize(10);

// Add some data, resembling some different data types
echo date('H:i:s') , " Add some data" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A1', 'String')
                              ->setCellValue('B1', 'Simple')
                              ->setCellValue('C1', '\PhpOffice\PhpSpreadsheet\Spreadsheet');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A2', 'String')
                              ->setCellValue('B2', 'Symbols')
                              ->setCellValue('C2', '!+&=()~§±æþ');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A3', 'String')
                              ->setCellValue('B3', 'UTF-8')
                              ->setCellValue('C3', 'Создать MS Excel Книги из PHP скриптов');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A4', 'Number')
                              ->setCellValue('B4', 'Integer')
                              ->setCellValue('C4', 12);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A5', 'Number')
                              ->setCellValue('B5', 'Float')
                              ->setCellValue('C5', 34.56);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A6', 'Number')
                              ->setCellValue('B6', 'Negative')
                              ->setCellValue('C6', -7.89);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A7', 'Boolean')
                              ->setCellValue('B7', 'True')
                              ->setCellValue('C7', true);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A8', 'Boolean')
                              ->setCellValue('B8', 'False')
                              ->setCellValue('C8', false);

$dateTimeNow = time();
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A9', 'Date/Time')
                              ->setCellValue('B9', 'Date')
                              ->setCellValue('C9', \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( $dateTimeNow ));
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('C9')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDD2);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A10', 'Date/Time')
                              ->setCellValue('B10', 'Time')
                              ->setCellValue('C10', \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( $dateTimeNow ));
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('C10')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME4);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A11', 'Date/Time')
                              ->setCellValue('B11', 'Date and Time')
                              ->setCellValue('C11', \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( $dateTimeNow ));
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('C11')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DATETIME);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A12', 'NULL')
                              ->setCellValue('C12', NULL);

$objRichText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
$objRichText->createText('你好 ');
$objPayable = $objRichText->createTextRun('你 好 吗？');
$objPayable->getFont()->setBold(true);
$objPayable->getFont()->setItalic(true);
$objPayable->getFont()->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_DARKGREEN ) );

$objRichText->createText(', unless specified otherwise on the invoice.');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A13', 'Rich Text')
                              ->setCellValue('C13', $objRichText);


$objRichText2 = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
$objRichText2->createText("black text\n");

$objRed = $objRichText2->createTextRun("red text");
$objRed->getFont()->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED  ) );

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getCell("C14")->setValue($objRichText2);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle("C14")->getAlignment()->setWrapText(true);


$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

// Rename worksheet
echo date('H:i:s') , " Rename worksheet" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setTitle('Datatypes');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0);


// Save Excel 95 file
echo date('H:i:s') , " Write to Excel5 format" , EOL;
$callStartTime = microtime(true);

$objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($obj\PhpOffice\PhpSpreadsheet\Spreadsheet, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xls', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


echo date('H:i:s') , " Reload workbook from saved file" , EOL;
$callStartTime = microtime(true);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(str_replace('.php', '.xls', __FILE__));

$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;
echo 'Call time to reload Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


var_dump($obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->toArray());


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done testing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;

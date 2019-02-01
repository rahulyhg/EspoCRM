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


// Create a first sheet, representing sales data
echo date('H:i:s') , " Add some data" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A1', 'Description')
                              ->setCellValue('B1', 'Amount');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A2', 'Paycheck received')
                              ->setCellValue('B2', 100);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A3', 'Cup of coffee bought')
                              ->setCellValue('B3', -1.5);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A4', 'Cup of coffee bought')
                              ->setCellValue('B4', -1.5);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A5', 'Cup of tea bought')
                              ->setCellValue('B5', -1.2);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A6', 'Found some money')
                              ->setCellValue('B6', 8);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A7', 'Total:')
                              ->setCellValue('B7', '=SUM(B2:B6)');


// Set column widths
echo date('H:i:s') , " Set column widths" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(12);


// Add conditional formatting
echo date('H:i:s') , " Add conditional formatting" , EOL;
$objConditional1 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$objConditional1->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS)
                ->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_BETWEEN)
                ->addCondition('200')
                ->addCondition('400');
$objConditional1->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_YELLOW);
$objConditional1->getStyle()->getFont()->setBold(true);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);

$objConditional2 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$objConditional2->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS)
                ->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_LESSTHAN)
                ->addCondition('0');
$objConditional2->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
$objConditional2->getStyle()->getFont()->setItalic(true);
$objConditional2->getStyle()->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);

$objConditional3 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$objConditional3->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS)
                ->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHANOREQUAL)
                ->addCondition('0');
$objConditional3->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);
$objConditional3->getStyle()->getFont()->setItalic(true);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);

$conditionalStyles = $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('B2')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional2);
array_push($conditionalStyles, $objConditional3);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('B2')->setConditionalStyles($conditionalStyles);


//	duplicate the conditional styles across a range of cells
echo date('H:i:s') , " Duplicate the conditional formatting across a range of cells" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->duplicateConditionalStyle(
				$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('B2')->getConditionalStyles(),
				'B3:B7'
			  );


// Set fonts
echo date('H:i:s') , " Set fonts" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('A1:B1')->getFont()->setBold(true);
//$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('A7:B7')->getFont()->setBold(true);
//$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('B7')->getFont()->setBold(true);


// Set header and footer. When no different headers for odd/even are used, odd header is assumed.
echo date('H:i:s') , " Set header/footer" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getHeaderFooter()->setOddHeader('&L&BPersonal cash register&RPrinted on &D');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' . $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->getTitle() . '&RPage &P of &N');


// Set page orientation and size
echo date('H:i:s') , " Set page orientation and size" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);


// Rename worksheet
echo date('H:i:s') , " Rename worksheet" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setTitle('Invoice');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0);


// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($obj\PhpOffice\PhpSpreadsheet\Spreadsheet, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Save Excel5 file
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

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
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()
    ->setCellValue('A1', '-0.5')
    ->setCellValue('A2', '-0.25')
    ->setCellValue('A3', '0.0')
    ->setCellValue('A4', '0.25')
    ->setCellValue('A5', '0.5')
    ->setCellValue('A6', '0.75')
    ->setCellValue('A7', '1.0')
    ->setCellValue('A8', '1.25')
;

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('A1:A8')
    ->getNumberFormat()
    ->setFormatCode(
        \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00
    );


// Add conditional formatting
echo date('H:i:s') , " Add conditional formatting" , EOL;
$objConditional1 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$objConditional1->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS)
    ->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);

$objConditional3 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$objConditional3->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS)
    ->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('1');
$objConditional3->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);

$conditionalStyles = $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('A1')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('A1')->setConditionalStyles($conditionalStyles);


//	duplicate the conditional styles across a range of cells
echo date('H:i:s') , " Duplicate the conditional formatting across a range of cells" , EOL;
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->duplicateConditionalStyle(
				$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getStyle('A1')->getConditionalStyles(),
				'A2:A8'
			  );


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

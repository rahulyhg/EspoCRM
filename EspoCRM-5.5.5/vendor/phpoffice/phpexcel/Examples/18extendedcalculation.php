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

date_default_timezone_set('Europe/London');

/** \PhpOffice\PhpSpreadsheet\Spreadsheet */
require_once dirname(__FILE__) . '/../Classes/\PhpOffice\PhpSpreadsheet\Spreadsheet.php';


// List functions
echo date('H:i:s') . " List implemented functions\n";
$objCalc = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance();
print_r($objCalc->listFunctionNames());

// Create new \PhpOffice\PhpSpreadsheet\Spreadsheet object
echo date('H:i:s') . " Create new \PhpOffice\PhpSpreadsheet\Spreadsheet object\n";
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();

// Add some data, we will use some formulas here
echo date('H:i:s') . " Add some data\n";
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('A14', 'Count:');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B1', 'Range 1');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B2', 2);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B3', 8);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B4', 10);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B5', True);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B6', False);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B7', 'Text String');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B9', '22');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B10', 4);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B11', 6);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B12', 12);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('B14', '=COUNT(B2:B12)');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C1', 'Range 2');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C2', 1);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C3', 2);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C4', 2);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C5', 3);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C6', 3);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C7', 3);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C8', '0');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C9', 4);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C10', 4);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C11', 4);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C12', 4);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('C14', '=COUNT(C2:C12)');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('D1', 'Range 3');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('D2', 2);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('D3', 3);
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('D4', 4);

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('D5', '=((D2 * D3) + D4) & " should be 10"');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('E1', 'Other functions');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('E2', '=PI()');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('E3', '=RAND()');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('E4', '=RANDBETWEEN(5, 10)');

$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('E14', 'Count of both ranges:');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setCellValue('F14', '=COUNT(B2:C12)');

// Calculated data
echo date('H:i:s') . " Calculated data\n";
echo 'Value of B14 [=COUNT(B2:B12)]: ' . $obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->getCell('B14')->getCalculatedValue() . "\r\n";


// Echo memory peak usage
echo date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB\r\n";

// Echo done
echo date('H:i:s') . " Done" , EOL;

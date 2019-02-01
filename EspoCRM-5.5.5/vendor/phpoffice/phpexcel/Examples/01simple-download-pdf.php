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

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include \PhpOffice\PhpSpreadsheet\Spreadsheet */
require_once dirname(__FILE__) . '/../Classes/\PhpOffice\PhpSpreadsheet\Spreadsheet.php';


//	Change these values to select the Rendering library that you wish to use
//		and its directory location on your server
//$rendererName = \PhpOffice\PhpSpreadsheet\Settings::PDF_RENDERER_TCPDF;
$rendererName = \PhpOffice\PhpSpreadsheet\Settings::PDF_RENDERER_MPDF;
//$rendererName = \PhpOffice\PhpSpreadsheet\Settings::PDF_RENDERER_DOMPDF;
//$rendererLibrary = 'tcPDF5.9';
$rendererLibrary = 'mPDF5.4';
//$rendererLibrary = 'domPDF0.6.0beta3';
$rendererLibraryPath = dirname(__FILE__).'/../../../libraries/PDF/' . $rendererLibrary;


// Create new \PhpOffice\PhpSpreadsheet\Spreadsheet object
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();

// Set document properties
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("PDF Test Document")
							 ->setSubject("PDF Test Document")
							 ->setDescription("Test document for PDF, generated using PHP classes.")
							 ->setKeywords("pdf php")
							 ->setCategory("Test result file");


// Add some data
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Hello')
            ->setCellValue('B2', 'world!')
            ->setCellValue('C1', 'Hello')
            ->setCellValue('D2', 'world!');

// Miscellaneous glyphs, UTF-8
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A4', 'Miscellaneous glyphs')
            ->setCellValue('A5', 'éàèùâêîôûëïüÿäöüç');

// Rename worksheet
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setTitle('Simple');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->getActiveSheet()->setShowGridLines(false);

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet->setActiveSheetIndex(0);


if (!\PhpOffice\PhpSpreadsheet\Settings::setPdfRenderer(
		$rendererName,
		$rendererLibraryPath
	)) {
	die(
		'NOTICE: Please set the $rendererName and $rendererLibraryPath values' .
		'<br />' .
		'at the top of this script as appropriate for your directory structure'
	);
}


// Redirect output to a client’s web browser (PDF)
header('Content-Type: application/pdf');
header('Content-Disposition: attachment;filename="01simple.pdf"');
header('Cache-Control: max-age=0');

$objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($obj\PhpOffice\PhpSpreadsheet\Spreadsheet, 'PDF');
$objWriter->save('php://output');
exit;

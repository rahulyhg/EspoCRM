<?php

// Turn off error reporting
error_reporting(0);

use PhpOffice\PhpSpreadsheet\IOFactory;

require __DIR__ . '/../Header.php';

$html = __DIR__ . '/../templates/46readHtml.html';
$callStartTime = microtime(true);

$objReader = IOFactory::createReader('Html');
$obj\PhpOffice\PhpSpreadsheet\Spreadsheet = $objReader->load($html);

$helper->logRead('Html', $html, $callStartTime);

// Save
$helper->write($obj\PhpOffice\PhpSpreadsheet\Spreadsheet, __FILE__);

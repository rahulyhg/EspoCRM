<?php
/**
 *  \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 *  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public
 *  License as published by the Free Software Foundation; either
 *  version 2.1 of the License, or (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  License along with this library; if not, write to the Free Software
 *  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *  @category    \PhpOffice\PhpSpreadsheet\Spreadsheet
 *  @package     \PhpOffice\PhpSpreadsheet\Writer\Pdf
 *  @copyright   Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 *  @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 *  @version     ##VERSION##, ##DATE##
 */


/**  Require tcPDF library */
$pdfRendererClassFile = \PhpOffice\PhpSpreadsheet\Settings::getPdfRendererPath() . '/tcpdf.php';
if (file_exists($pdfRendererClassFile)) {
    $k_path_url = \PhpOffice\PhpSpreadsheet\Settings::getPdfRendererPath();
    require_once $pdfRendererClassFile;
} else {
    throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Unable to load PDF Rendering library');
}

/**
 *  \PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf
 *
 *  @category    \PhpOffice\PhpSpreadsheet\Spreadsheet
 *  @package     \PhpOffice\PhpSpreadsheet\Writer\Pdf
 *  @copyright   Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf extends \PhpOffice\PhpSpreadsheet\Writer\Pdf implements \PhpOffice\PhpSpreadsheet\Writer\IWriter
{
    /**
     *  Create a new \PhpOffice\PhpSpreadsheet\Writer\Pdf
     *
     *  @param  \PhpOffice\PhpSpreadsheet\Spreadsheet  $phpExcel  \PhpOffice\PhpSpreadsheet\Spreadsheet object
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel)
    {
        parent::__construct($phpExcel);
    }

    /**
     *  Save \PhpOffice\PhpSpreadsheet\Spreadsheet to file
     *
     *  @param     string     $pFilename   Name of the file to save as
     *  @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save($pFilename = NULL)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'LETTER';    //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (is_null($this->getSheetIndex())) {
            $orientation = ($this->_phpExcel->getSheet(0)->getPageSetup()->getOrientation()
                == \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_phpExcel->getSheet(0)->getPageSetup()->getPaperSize();
            $printMargins = $this->_phpExcel->getSheet(0)->getPageMargins();
        } else {
            $orientation = ($this->_phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
            $printMargins = $this->_phpExcel->getSheet($this->getSheetIndex())->getPageMargins();
        }

        //  Override Page Orientation
        if (!is_null($this->getOrientation())) {
            $orientation = ($this->getOrientation() == \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE)
                ? 'L'
                : 'P';
        }
        //  Override Paper Size
        if (!is_null($this->getPaperSize())) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$_paperSizes[$printPaperSize])) {
            $paperSize = self::$_paperSizes[$printPaperSize];
        }


        //  Create PDF
        $pdf = new TCPDF($orientation, 'pt', $paperSize);
        $pdf->setFontSubsetting(FALSE);
        //    Set margins, converting inches to points (using 72 dpi)
        $pdf->SetMargins($printMargins->getLeft() * 72, $printMargins->getTop() * 72, $printMargins->getRight() * 72);
        $pdf->SetAutoPageBreak(TRUE, $printMargins->getBottom() * 72);

        $pdf->setPrintHeader(FALSE);
        $pdf->setPrintFooter(FALSE);

        $pdf->AddPage();

        //  Set the appropriate font
        $pdf->SetFont($this->getFont());
        $pdf->writeHTML(
            $this->generateHTMLHeader(FALSE) .
            $this->generateSheetData() .
            $this->generateHTMLFooter()
        );

        //  Document info
        $pdf->SetTitle($this->_phpExcel->getProperties()->getTitle());
        $pdf->SetAuthor($this->_phpExcel->getProperties()->getCreator());
        $pdf->SetSubject($this->_phpExcel->getProperties()->getSubject());
        $pdf->SetKeywords($this->_phpExcel->getProperties()->getKeywords());
        $pdf->SetCreator($this->_phpExcel->getProperties()->getCreator());

        //  Write to file
        fwrite($fileHandle, $pdf->output($pFilename, 'S'));

		parent::restoreStateAfterSave($fileHandle);
    }

}

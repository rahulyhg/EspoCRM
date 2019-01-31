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
 * @package    \PhpOffice\PhpSpreadsheet\Writer\Ods
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Writer\Ods
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Writer\Ods
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @author     Alexander Pervakov <frost-nzcr4@jagmort.com>
 * @link       http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os.html
 */
class \PhpOffice\PhpSpreadsheet\Writer\Ods extends \PhpOffice\PhpSpreadsheet\Writer\BaseWriter implements \PhpOffice\PhpSpreadsheet\Writer\IWriter
{
    /**
     * Private writer parts
     *
     * @var \PhpOffice\PhpSpreadsheet\Writer\Ods\WriterPart[]
     */
    private $_writerParts = array();

    /**
     * Private \PhpOffice\PhpSpreadsheet\Spreadsheet
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    private $_spreadSheet;

    /**
     * Create a new \PhpOffice\PhpSpreadsheet\Writer\Ods
     *
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
    {
        $this->set\PhpOffice\PhpSpreadsheet\Spreadsheet($p\PhpOffice\PhpSpreadsheet\Spreadsheet);

        $writerPartsArray = array(
            'content'    => '\PhpOffice\PhpSpreadsheet\Writer\Ods\Content',
            'meta'       => '\PhpOffice\PhpSpreadsheet\Writer\Ods\Meta',
            'meta_inf'   => '\PhpOffice\PhpSpreadsheet\Writer\Ods\MetaInf',
            'mimetype'   => '\PhpOffice\PhpSpreadsheet\Writer\Ods\Mimetype',
            'settings'   => '\PhpOffice\PhpSpreadsheet\Writer\Ods\Settings',
            'styles'     => '\PhpOffice\PhpSpreadsheet\Writer\Ods\Styles',
            'thumbnails' => '\PhpOffice\PhpSpreadsheet\Writer\Ods\Thumbnails'
        );

        foreach ($writerPartsArray as $writer => $class) {
            $this->_writerParts[$writer] = new $class($this);
        }
    }

    /**
     * Get writer part
     *
     * @param  string  $pPartName  Writer part name
     * @return \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
     */
    public function getWriterPart($pPartName = '')
    {
        if ($pPartName != '' && isset($this->_writerParts[strtolower($pPartName)])) {
            return $this->_writerParts[strtolower($pPartName)];
        } else {
            return null;
        }
    }

    /**
     * Save \PhpOffice\PhpSpreadsheet\Spreadsheet to file
     *
     * @param  string  $pFilename
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save($pFilename = NULL)
    {
        if (!$this->_spreadSheet) {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('\PhpOffice\PhpSpreadsheet\Spreadsheet object unassigned.');
        }

        // garbage collect
        $this->_spreadSheet->garbageCollect();

        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if (strtolower($pFilename) == 'php://output' || strtolower($pFilename) == 'php://stdout') {
            $pFilename = @tempnam(\PhpOffice\PhpSpreadsheet\Shared\File::sys_get_temp_dir(), 'phpxltmp');
            if ($pFilename == '') {
                $pFilename = $originalFilename;
            }
        }

        $objZip = $this->_createZip($pFilename);

        $objZip->addFromString('META-INF/manifest.xml', $this->getWriterPart('meta_inf')->writeManifest());
        $objZip->addFromString('Thumbnails/thumbnail.png', $this->getWriterPart('thumbnails')->writeThumbnail());
        $objZip->addFromString('content.xml',  $this->getWriterPart('content')->write());
        $objZip->addFromString('meta.xml',     $this->getWriterPart('meta')->write());
        $objZip->addFromString('mimetype',     $this->getWriterPart('mimetype')->write());
        $objZip->addFromString('settings.xml', $this->getWriterPart('settings')->write());
        $objZip->addFromString('styles.xml',   $this->getWriterPart('styles')->write());

        // Close file
        if ($objZip->close() === false) {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Could not close zip file $pFilename.");
        }

        // If a temporary file was used, copy it to the correct file stream
        if ($originalFilename != $pFilename) {
            if (copy($pFilename, $originalFilename) === false) {
                throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Could not copy temporary zip file $pFilename to $originalFilename.");
            }
            @unlink($pFilename);
        }
    }

    /**
     * Create zip object
     *
     * @param string $pFilename
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @return ZipArchive
     */
    private function _createZip($pFilename)
    {
        // Create new ZIP file and open it for writing
        $zipClass = \PhpOffice\PhpSpreadsheet\Settings::getZipClass();
        $objZip = new $zipClass();

        // Retrieve OVERWRITE and CREATE constants from the instantiated zip class
        // This method of accessing constant values from a dynamic class should work with all appropriate versions of PHP
        $ro = new ReflectionObject($objZip);
        $zipOverWrite = $ro->getConstant('OVERWRITE');
        $zipCreate = $ro->getConstant('CREATE');

        if (file_exists($pFilename)) {
            unlink($pFilename);
        }
        // Try opening the ZIP file
        if ($objZip->open($pFilename, $zipOverWrite) !== true) {
            if ($objZip->open($pFilename, $zipCreate) !== true) {
                throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Could not open $pFilename for writing.");
            }
        }

        return $objZip;
    }

    /**
     * Get \PhpOffice\PhpSpreadsheet\Spreadsheet object
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function get\PhpOffice\PhpSpreadsheet\Spreadsheet()
    {
        if ($this->_spreadSheet !== null) {
            return $this->_spreadSheet;
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('No \PhpOffice\PhpSpreadsheet\Spreadsheet assigned.');
        }
    }

    /**
     * Set \PhpOffice\PhpSpreadsheet\Spreadsheet object
     *
     * @param  \PhpOffice\PhpSpreadsheet\Spreadsheet  $p\PhpOffice\PhpSpreadsheet\Spreadsheet  \PhpOffice\PhpSpreadsheet\Spreadsheet object
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @return \PhpOffice\PhpSpreadsheet\Writer\Xlsx
     */
    public function set\PhpOffice\PhpSpreadsheet\Spreadsheet(\PhpOffice\PhpSpreadsheet\Spreadsheet $p\PhpOffice\PhpSpreadsheet\Spreadsheet = null)
    {
        $this->_spreadSheet = $p\PhpOffice\PhpSpreadsheet\Spreadsheet;
        return $this;
    }
}

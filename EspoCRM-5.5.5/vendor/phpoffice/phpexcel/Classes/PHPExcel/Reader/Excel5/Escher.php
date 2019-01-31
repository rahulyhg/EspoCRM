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
 * @package    \PhpOffice\PhpSpreadsheet\Reader\Xls
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Reader\Xls
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher
{
	const DGGCONTAINER		= 0xF000;
	const BSTORECONTAINER	= 0xF001;
	const DGCONTAINER		= 0xF002;
	const SPGRCONTAINER		= 0xF003;
	const SPCONTAINER		= 0xF004;
	const DGG				= 0xF006;
	const BSE				= 0xF007;
	const DG				= 0xF008;
	const SPGR				= 0xF009;
	const SP				= 0xF00A;
	const OPT				= 0xF00B;
	const CLIENTTEXTBOX		= 0xF00D;
	const CLIENTANCHOR		= 0xF010;
	const CLIENTDATA		= 0xF011;
	const BLIPJPEG			= 0xF01D;
	const BLIPPNG			= 0xF01E;
	const SPLITMENUCOLORS	= 0xF11E;
	const TERTIARYOPT		= 0xF122;

	/**
	 * Escher stream data (binary)
	 *
	 * @var string
	 */
	private $_data;

	/**
	 * Size in bytes of the Escher stream data
	 *
	 * @var int
	 */
	private $_dataSize;

	/**
	 * Current position of stream pointer in Escher stream data
	 *
	 * @var int
	 */
	private $_pos;

	/**
	 * The object to be returned by the reader. Modified during load.
	 *
	 * @var mixed
	 */
	private $_object;

	/**
	 * Create a new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher instance
	 *
	 * @param mixed $object
	 */
	public function __construct($object)
	{
		$this->_object = $object;
	}

	/**
	 * Load Escher stream data. May be a partial Escher stream.
	 *
	 * @param string $data
	 */
	public function load($data)
	{
		$this->_data = $data;

		// total byte size of Excel data (workbook global substream + sheet substreams)
		$this->_dataSize = strlen($this->_data);

		$this->_pos = 0;

		// Parse Escher stream
		while ($this->_pos < $this->_dataSize) {

			// offset: 2; size: 2: Record Type
			$fbt = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos + 2);

			switch ($fbt) {
				case self::DGGCONTAINER:	$this->_readDggContainer();		break;
				case self::DGG:				$this->_readDgg();				break;
				case self::BSTORECONTAINER:	$this->_readBstoreContainer();	break;
				case self::BSE:				$this->_readBSE();				break;
				case self::BLIPJPEG:		$this->_readBlipJPEG();			break;
				case self::BLIPPNG:			$this->_readBlipPNG();			break;
				case self::OPT:				$this->_readOPT();				break;
				case self::TERTIARYOPT:		$this->_readTertiaryOPT();		break;
				case self::SPLITMENUCOLORS:	$this->_readSplitMenuColors();	break;
				case self::DGCONTAINER:		$this->_readDgContainer();		break;
				case self::DG:				$this->_readDg();				break;
				case self::SPGRCONTAINER:	$this->_readSpgrContainer();	break;
				case self::SPCONTAINER:		$this->_readSpContainer();		break;
				case self::SPGR:			$this->_readSpgr();				break;
				case self::SP:				$this->_readSp();				break;
				case self::CLIENTTEXTBOX:	$this->_readClientTextbox();	break;
				case self::CLIENTANCHOR:	$this->_readClientAnchor();		break;
				case self::CLIENTDATA:		$this->_readClientData();		break;
				default:					$this->_readDefault();			break;
			}
		}

		return $this->_object;
	}

	/**
	 * Read a generic record
	 */
	private function _readDefault()
	{
		// offset 0; size: 2; recVer and recInstance
		$verInstance = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos);

		// offset: 2; size: 2: Record Type
		$fbt = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos + 2);

		// bit: 0-3; mask: 0x000F; recVer
		$recVer = (0x000F & $verInstance) >> 0;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read DggContainer record (Drawing Group Container)
	 */
	private function _readDggContainer()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// record is a container, read contents
		$dggContainer = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer();
		$this->_object->setDggContainer($dggContainer);
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher($dggContainer);
		$reader->load($recordData);
	}

	/**
	 * Read Dgg record (Drawing Group)
	 */
	private function _readDgg()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read BstoreContainer record (Blip Store Container)
	 */
	private function _readBstoreContainer()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// record is a container, read contents
		$bstoreContainer = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer();
		$this->_object->setBstoreContainer($bstoreContainer);
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher($bstoreContainer);
		$reader->load($recordData);
	}

	/**
	 * Read BSE record
	 */
	private function _readBSE()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// add BSE to BstoreContainer
		$BSE = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer\BSE();
		$this->_object->addBSE($BSE);

		$BSE->setBLIPType($recInstance);

		// offset: 0; size: 1; btWin32 (MSOBLIPTYPE)
		$btWin32 = ord($recordData[0]);

		// offset: 1; size: 1; btWin32 (MSOBLIPTYPE)
		$btMacOS = ord($recordData[1]);

		// offset: 2; size: 16; MD4 digest
		$rgbUid = substr($recordData, 2, 16);

		// offset: 18; size: 2; tag
		$tag = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 18);

		// offset: 20; size: 4; size of BLIP in bytes
		$size = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($recordData, 20);

		// offset: 24; size: 4; number of references to this BLIP
		$cRef = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($recordData, 24);

		// offset: 28; size: 4; MSOFO file offset
		$foDelay = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($recordData, 28);

		// offset: 32; size: 1; unused1
		$unused1 = ord($recordData{32});

		// offset: 33; size: 1; size of nameData in bytes (including null terminator)
		$cbName = ord($recordData{33});

		// offset: 34; size: 1; unused2
		$unused2 = ord($recordData{34});

		// offset: 35; size: 1; unused3
		$unused3 = ord($recordData{35});

		// offset: 36; size: $cbName; nameData
		$nameData = substr($recordData, 36, $cbName);

		// offset: 36 + $cbName, size: var; the BLIP data
		$blipData = substr($recordData, 36 + $cbName);

		// record is a container, read contents
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher($BSE);
		$reader->load($blipData);
	}

	/**
	 * Read BlipJPEG record. Holds raw JPEG image data
	 */
	private function _readBlipJPEG()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		$pos = 0;

		// offset: 0; size: 16; rgbUid1 (MD4 digest of)
		$rgbUid1 = substr($recordData, 0, 16);
		$pos += 16;

		// offset: 16; size: 16; rgbUid2 (MD4 digest), only if $recInstance = 0x46B or 0x6E3
		if (in_array($recInstance, array(0x046B, 0x06E3))) {
			$rgbUid2 = substr($recordData, 16, 16);
			$pos += 16;
		}

		// offset: var; size: 1; tag
		$tag = ord($recordData{$pos});
		$pos += 1;

		// offset: var; size: var; the raw image data
		$data = substr($recordData, $pos);

		$blip = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer\BSE\Blip();
		$blip->setData($data);

		$this->_object->setBlip($blip);
	}

	/**
	 * Read BlipPNG record. Holds raw PNG image data
	 */
	private function _readBlipPNG()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		$pos = 0;

		// offset: 0; size: 16; rgbUid1 (MD4 digest of)
		$rgbUid1 = substr($recordData, 0, 16);
		$pos += 16;

		// offset: 16; size: 16; rgbUid2 (MD4 digest), only if $recInstance = 0x46B or 0x6E3
		if ($recInstance == 0x06E1) {
			$rgbUid2 = substr($recordData, 16, 16);
			$pos += 16;
		}

		// offset: var; size: 1; tag
		$tag = ord($recordData{$pos});
		$pos += 1;

		// offset: var; size: var; the raw image data
		$data = substr($recordData, $pos);

		$blip = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer\BSE\Blip();
		$blip->setData($data);

		$this->_object->setBlip($blip);
	}

	/**
	 * Read OPT record. This record may occur within DggContainer record or SpContainer
	 */
	private function _readOPT()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		$this->_readOfficeArtRGFOPTE($recordData, $recInstance);
	}

	/**
	 * Read TertiaryOPT record
	 */
	private function _readTertiaryOPT()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read SplitMenuColors record
	 */
	private function _readSplitMenuColors()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read DgContainer record (Drawing Container)
	 */
	private function _readDgContainer()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// record is a container, read contents
		$dgContainer = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer();
		$this->_object->setDgContainer($dgContainer);
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher($dgContainer);
		$escher = $reader->load($recordData);
	}

	/**
	 * Read Dg record (Drawing)
	 */
	private function _readDg()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read SpgrContainer record (Shape Group Container)
	 */
	private function _readSpgrContainer()
	{
		// context is either context DgContainer or SpgrContainer

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// record is a container, read contents
		$spgrContainer = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer\SpgrContainer();

		if ($this->_object instanceof \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer) {
			// DgContainer
			$this->_object->setSpgrContainer($spgrContainer);
		} else {
			// SpgrContainer
			$this->_object->addChild($spgrContainer);
		}

		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher($spgrContainer);
		$escher = $reader->load($recordData);
	}

	/**
	 * Read SpContainer record (Shape Container)
	 */
	private function _readSpContainer()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// add spContainer to spgrContainer
		$spContainer = new \PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer\SpgrContainer\SpContainer();
		$this->_object->addChild($spContainer);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// record is a container, read contents
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls\Escher($spContainer);
		$escher = $reader->load($recordData);
	}

	/**
	 * Read Spgr record (Shape Group)
	 */
	private function _readSpgr()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read Sp record (Shape)
	 */
	private function _readSp()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read ClientTextbox record
	 */
	private function _readClientTextbox()
	{
		// offset: 0; size: 2; recVer and recInstance

		// bit: 4-15; mask: 0xFFF0; recInstance
		$recInstance = (0xFFF0 & \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($this->_data, $this->_pos)) >> 4;

		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read ClientAnchor record. This record holds information about where the shape is anchored in worksheet
	 */
	private function _readClientAnchor()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;

		// offset: 2; size: 2; upper-left corner column index (0-based)
		$c1 = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 2);

		// offset: 4; size: 2; upper-left corner horizontal offset in 1/1024 of column width
		$startOffsetX = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 4);

		// offset: 6; size: 2; upper-left corner row index (0-based)
		$r1 = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 6);

		// offset: 8; size: 2; upper-left corner vertical offset in 1/256 of row height
		$startOffsetY = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 8);

		// offset: 10; size: 2; bottom-right corner column index (0-based)
		$c2 = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 10);

		// offset: 12; size: 2; bottom-right corner horizontal offset in 1/1024 of column width
		$endOffsetX = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 12);

		// offset: 14; size: 2; bottom-right corner row index (0-based)
		$r2 = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 14);

		// offset: 16; size: 2; bottom-right corner vertical offset in 1/256 of row height
		$endOffsetY = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($recordData, 16);

		// set the start coordinates
		$this->_object->setStartCoordinates(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($c1) . ($r1 + 1));

		// set the start offsetX
		$this->_object->setStartOffsetX($startOffsetX);

		// set the start offsetY
		$this->_object->setStartOffsetY($startOffsetY);

		// set the end coordinates
		$this->_object->setEndCoordinates(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($c2) . ($r2 + 1));

		// set the end offsetX
		$this->_object->setEndOffsetX($endOffsetX);

		// set the end offsetY
		$this->_object->setEndOffsetY($endOffsetY);
	}

	/**
	 * Read ClientData record
	 */
	private function _readClientData()
	{
		$length = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($this->_data, $this->_pos + 4);
		$recordData = substr($this->_data, $this->_pos + 8, $length);

		// move stream pointer to next record
		$this->_pos += 8 + $length;
	}

	/**
	 * Read OfficeArtRGFOPTE table of property-value pairs
	 *
	 * @param string $data Binary data
	 * @param int $n Number of properties
	 */
	private function _readOfficeArtRGFOPTE($data, $n) {

		$splicedComplexData = substr($data, 6 * $n);

		// loop through property-value pairs
		for ($i = 0; $i < $n; ++$i) {
			// read 6 bytes at a time
			$fopte = substr($data, 6 * $i, 6);

			// offset: 0; size: 2; opid
			$opid = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt2d($fopte, 0);

			// bit: 0-13; mask: 0x3FFF; opid.opid
			$opidOpid = (0x3FFF & $opid) >> 0;

			// bit: 14; mask 0x4000; 1 = value in op field is BLIP identifier
			$opidFBid = (0x4000 & $opid) >> 14;

			// bit: 15; mask 0x8000; 1 = this is a complex property, op field specifies size of complex data
			$opidFComplex = (0x8000 & $opid) >> 15;

			// offset: 2; size: 4; the value for this property
			$op = \PhpOffice\PhpSpreadsheet\Reader\Xls::_GetInt4d($fopte, 2);

			if ($opidFComplex) {
				$complexData = substr($splicedComplexData, 0, $op);
				$splicedComplexData = substr($splicedComplexData, $op);

				// we store string value with complex data
				$value = $complexData;
			} else {
				// we store integer value
				$value = $op;
			}

			$this->_object->setOPT($opidOpid, $value);
		}
	}

}

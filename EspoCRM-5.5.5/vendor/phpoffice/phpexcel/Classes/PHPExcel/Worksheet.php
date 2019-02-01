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
 * @package    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 *
 * @category   \PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
 * @copyright  Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet implements \PhpOffice\PhpSpreadsheet\IComparable
{
    /* Break types */
    const BREAK_NONE   = 0;
    const BREAK_ROW    = 1;
    const BREAK_COLUMN = 2;

    /* Sheet state */
    const SHEETSTATE_VISIBLE    = 'visible';
    const SHEETSTATE_HIDDEN     = 'hidden';
    const SHEETSTATE_VERYHIDDEN = 'veryHidden';

    /**
     * Invalid characters in sheet title
     *
     * @var array
     */
    private static $_invalidCharacters = array('*', ':', '/', '\\', '?', '[', ']');

    /**
     * Parent spreadsheet
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    private $_parent;

    /**
     * Cacheable collection of cells
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet_CachedObjectStorage_xxx
     */
    private $_cellCollection = null;

    /**
     * Collection of row dimensions
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension[]
     */
    private $_rowDimensions = array();

    /**
     * Default row dimension
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension
     */
    private $_defaultRowDimension = null;

    /**
     * Collection of column dimensions
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension[]
     */
    private $_columnDimensions = array();

    /**
     * Default column dimension
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension
     */
    private $_defaultColumnDimension = null;

    /**
     * Collection of drawings
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\BaseDrawing[]
     */
    private $_drawingCollection = null;

    /**
     * Collection of Chart objects
     *
     * @var \PhpOffice\PhpSpreadsheet\Chart\Chart[]
     */
    private $_chartCollection = array();

    /**
     * Worksheet title
     *
     * @var string
     */
    private $_title;

    /**
     * Sheet state
     *
     * @var string
     */
    private $_sheetState;

    /**
     * Page setup
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup
     */
    private $_pageSetup;

    /**
     * Page margins
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\PageMargins
     */
    private $_pageMargins;

    /**
     * Page header/footer
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter
     */
    private $_headerFooter;

    /**
     * Sheet view
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\SheetView
     */
    private $_sheetView;

    /**
     * Protection
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    private $_protection;

    /**
     * Collection of styles
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Style[]
     */
    private $_styles = array();

    /**
     * Conditional styles. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private $_conditionalStylesCollection = array();

    /**
     * Is the current cell collection sorted already?
     *
     * @var boolean
     */
    private $_cellCollectionIsSorted = false;

    /**
     * Collection of breaks
     *
     * @var array
     */
    private $_breaks = array();

    /**
     * Collection of merged cell ranges
     *
     * @var array
     */
    private $_mergeCells = array();

    /**
     * Collection of protected cell ranges
     *
     * @var array
     */
    private $_protectedCells = array();

    /**
     * Autofilter Range and selection
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    private $_autoFilter = NULL;

    /**
     * Freeze pane
     *
     * @var string
     */
    private $_freezePane = '';

    /**
     * Show gridlines?
     *
     * @var boolean
     */
    private $_showGridlines = true;

    /**
    * Print gridlines?
    *
    * @var boolean
    */
    private $_printGridlines = false;

    /**
    * Show row and column headers?
    *
    * @var boolean
    */
    private $_showRowColHeaders = true;

    /**
     * Show summary below? (Row/Column outline)
     *
     * @var boolean
     */
    private $_showSummaryBelow = true;

    /**
     * Show summary right? (Row/Column outline)
     *
     * @var boolean
     */
    private $_showSummaryRight = true;

    /**
     * Collection of comments
     *
     * @var \PhpOffice\PhpSpreadsheet\Comment[]
     */
    private $_comments = array();

    /**
     * Active cell. (Only one!)
     *
     * @var string
     */
    private $_activeCell = 'A1';

    /**
     * Selected cells
     *
     * @var string
     */
    private $_selectedCells = 'A1';

    /**
     * Cached highest column
     *
     * @var string
     */
    private $_cachedHighestColumn = 'A';

    /**
     * Cached highest row
     *
     * @var int
     */
    private $_cachedHighestRow = 1;

    /**
     * Right-to-left?
     *
     * @var boolean
     */
    private $_rightToLeft = false;

    /**
     * Hyperlinks. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private $_hyperlinkCollection = array();

    /**
     * Data validation objects. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private $_dataValidationCollection = array();

    /**
     * Tab color
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Color
     */
    private $_tabColor;

    /**
     * Dirty flag
     *
     * @var boolean
     */
    private $_dirty    = true;

    /**
     * Hash
     *
     * @var string
     */
    private $_hash    = null;

    /**
    * CodeName
    *
    * @var string
    */
    private $_codeName = null;

	/**
     * Create a new worksheet
     *
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet        $pParent
     * @param string        $pTitle
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $pParent = null, $pTitle = 'Worksheet')
    {
        // Set parent and title
        $this->_parent = $pParent;
        $this->setTitle($pTitle, FALSE);
        // setTitle can change $pTitle
	    $this->setCodeName($this->getTitle());
        $this->setSheetState(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::SHEETSTATE_VISIBLE);

        $this->_cellCollection        = \PhpOffice\PhpSpreadsheet\Collection\CellsFactory::getInstance($this);

        // Set page setup
        $this->_pageSetup            = new \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup();

        // Set page margins
        $this->_pageMargins         = new \PhpOffice\PhpSpreadsheet\Worksheet\PageMargins();

        // Set page header/footer
        $this->_headerFooter        = new \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter();

        // Set sheet view
        $this->_sheetView            = new \PhpOffice\PhpSpreadsheet\Worksheet\SheetView();

        // Drawing collection
        $this->_drawingCollection    = new ArrayObject();

        // Chart collection
        $this->_chartCollection     = new ArrayObject();

        // Protection
        $this->_protection            = new \PhpOffice\PhpSpreadsheet\Worksheet\Protection();

        // Default row dimension
        $this->_defaultRowDimension = new \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension(NULL);

        // Default column dimension
        $this->_defaultColumnDimension    = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension(NULL);

        $this->_autoFilter            = new \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter(NULL, $this);
    }


    /**
     * Disconnect all cells from this \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet object,
     *    typically so that the worksheet object can be unset
     *
     */
	public function disconnectCells() {
    	if ( $this->_cellCollection !== NULL){
            $this->_cellCollection->unsetWorksheetCells();
            $this->_cellCollection = NULL;
    	}
        //    detach ourself from the workbook, so that it can then delete this worksheet successfully
        $this->_parent = null;
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
	function __destruct() {
		\PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($this->_parent)
		    ->clearCalculationCacheForWorksheet($this->_title);

		$this->disconnectCells();
	}

   /**
     * Return the cache controller for the cell collection
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet_CachedObjectStorage_xxx
     */
	public function getCellCacheController() {
        return $this->_cellCollection;
    }    //    function getCellCacheController()


    /**
     * Get array of invalid characters for sheet title
     *
     * @return array
     */
    public static function getInvalidCharacters()
    {
        return self::$_invalidCharacters;
    }

    /**
     * Check sheet code name for valid Excel syntax
     *
     * @param string $pValue The string to check
     * @return string The valid string
     * @throws Exception
     */
    private static function _checkSheetCodeName($pValue)
    {
        $CharCount = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue);
        if ($CharCount == 0) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Sheet code name cannot be empty.');
        }
        // Some of the printable ASCII characters are invalid:  * : / \ ? [ ] and  first and last characters cannot be a "'"
        if ((str_replace(self::$_invalidCharacters, '', $pValue) !== $pValue) || 
            (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,-1,1)=='\'') || 
            (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,1)=='\'')) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Invalid character found in sheet code name');
        }
 
        // Maximum 31 characters allowed for sheet title
        if ($CharCount > 31) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Maximum 31 characters allowed in sheet code name.');
        }
 
        return $pValue;
    }

   /**
     * Check sheet title for valid Excel syntax
     *
     * @param string $pValue The string to check
     * @return string The valid string
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private static function _checkSheetTitle($pValue)
    {
        // Some of the printable ASCII characters are invalid:  * : / \ ? [ ]
        if (str_replace(self::$_invalidCharacters, '', $pValue) !== $pValue) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Invalid character found in sheet title');
        }

        // Maximum 31 characters allowed for sheet title
        if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 31) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Maximum 31 characters allowed in sheet title.');
        }

        return $pValue;
    }

    /**
     * Get collection of cells
     *
     * @param boolean $pSorted Also sort the cell collection?
     * @return \PhpOffice\PhpSpreadsheet\Cell\Cell[]
     */
    public function getCellCollection($pSorted = true)
    {
        if ($pSorted) {
            // Re-order cell collection
            return $this->sortCellCollection();
        }
        if ($this->_cellCollection !== NULL) {
            return $this->_cellCollection->getCellList();
        }
        return array();
    }

    /**
     * Sort collection of cells
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function sortCellCollection()
    {
        if ($this->_cellCollection !== NULL) {
            return $this->_cellCollection->getSortedCellList();
        }
        return array();
    }

    /**
     * Get collection of row dimensions
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension[]
     */
    public function getRowDimensions()
    {
        return $this->_rowDimensions;
    }

    /**
     * Get default row dimension
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension
     */
    public function getDefaultRowDimension()
    {
        return $this->_defaultRowDimension;
    }

    /**
     * Get collection of column dimensions
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension[]
     */
    public function getColumnDimensions()
    {
        return $this->_columnDimensions;
    }

    /**
     * Get default column dimension
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension
     */
    public function getDefaultColumnDimension()
    {
        return $this->_defaultColumnDimension;
    }

    /**
     * Get collection of drawings
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\BaseDrawing[]
     */
    public function getDrawingCollection()
    {
        return $this->_drawingCollection;
    }

    /**
     * Get collection of charts
     *
     * @return \PhpOffice\PhpSpreadsheet\Chart\Chart[]
     */
    public function getChartCollection()
    {
        return $this->_chartCollection;
    }

    /**
     * Add chart
     *
     * @param \PhpOffice\PhpSpreadsheet\Chart\Chart $pChart
     * @param int|null $iChartIndex Index where chart should go (0,1,..., or null for last)
     * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
     */
    public function addChart(\PhpOffice\PhpSpreadsheet\Chart\Chart $pChart = null, $iChartIndex = null)
    {
        $pChart->setWorksheet($this);
        if (is_null($iChartIndex)) {
            $this->_chartCollection[] = $pChart;
        } else {
            // Insert the chart at the requested index
            array_splice($this->_chartCollection, $iChartIndex, 0, array($pChart));
        }

        return $pChart;
    }

    /**
     * Return the count of charts on this worksheet
     *
     * @return int        The number of charts
     */
    public function getChartCount()
    {
        return count($this->_chartCollection);
    }

    /**
     * Get a chart by its index position
     *
     * @param string $index Chart index position
     * @return false|\PhpOffice\PhpSpreadsheet\Chart\Chart
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getChartByIndex($index = null)
    {
        $chartCount = count($this->_chartCollection);
        if ($chartCount == 0) {
            return false;
        }
        if (is_null($index)) {
            $index = --$chartCount;
        }
        if (!isset($this->_chartCollection[$index])) {
            return false;
        }

        return $this->_chartCollection[$index];
    }

    /**
     * Return an array of the names of charts on this worksheet
     *
     * @return string[] The names of charts
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getChartNames()
    {
        $chartNames = array();
        foreach($this->_chartCollection as $chart) {
            $chartNames[] = $chart->getName();
        }
        return $chartNames;
    }

    /**
     * Get a chart by name
     *
     * @param string $chartName Chart name
     * @return false|\PhpOffice\PhpSpreadsheet\Chart\Chart
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getChartByName($chartName = '')
    {
        $chartCount = count($this->_chartCollection);
        if ($chartCount == 0) {
            return false;
        }
        foreach($this->_chartCollection as $index => $chart) {
            if ($chart->getName() == $chartName) {
                return $this->_chartCollection[$index];
            }
        }
        return false;
    }

    /**
     * Refresh column dimensions
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function refreshColumnDimensions()
    {
        $currentColumnDimensions = $this->getColumnDimensions();
        $newColumnDimensions = array();

        foreach ($currentColumnDimensions as $objColumnDimension) {
            $newColumnDimensions[$objColumnDimension->getColumnIndex()] = $objColumnDimension;
        }

        $this->_columnDimensions = $newColumnDimensions;

        return $this;
    }

    /**
     * Refresh row dimensions
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function refreshRowDimensions()
    {
        $currentRowDimensions = $this->getRowDimensions();
        $newRowDimensions = array();

        foreach ($currentRowDimensions as $objRowDimension) {
            $newRowDimensions[$objRowDimension->getRowIndex()] = $objRowDimension;
        }

        $this->_rowDimensions = $newRowDimensions;

        return $this;
    }

    /**
     * Calculate worksheet dimension
     *
     * @return string  String containing the dimension of this worksheet
     */
    public function calculateWorksheetDimension()
    {
        // Return
        return 'A1' . ':' .  $this->getHighestColumn() . $this->getHighestRow();
    }

    /**
     * Calculate worksheet data dimension
     *
     * @return string  String containing the dimension of this worksheet that actually contain data
     */
    public function calculateWorksheetDataDimension()
    {
        // Return
        return 'A1' . ':' .  $this->getHighestDataColumn() . $this->getHighestDataRow();
    }

    /**
     * Calculate widths for auto-size columns
     *
     * @param  boolean  $calculateMergeCells  Calculate merge cell width
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
     */
    public function calculateColumnWidths($calculateMergeCells = false)
    {
        // initialize $autoSizes array
        $autoSizes = array();
        foreach ($this->getColumnDimensions() as $colDimension) {
            if ($colDimension->getAutoSize()) {
                $autoSizes[$colDimension->getColumnIndex()] = -1;
            }
        }

        // There is only something to do if there are some auto-size columns
        if (!empty($autoSizes)) {

            // build list of cells references that participate in a merge
            $isMergeCell = array();
            foreach ($this->getMergeCells() as $cells) {
                foreach (\PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($cells) as $cellReference) {
                    $isMergeCell[$cellReference] = true;
                }
            }

            // loop through all cells in the worksheet
            foreach ($this->getCellCollection(false) as $cellID) {
                $cell = $this->getCell($cellID);
				if (isset($autoSizes[$this->_cellCollection->getCurrentColumn()])) {
                    // Determine width if cell does not participate in a merge
					if (!isset($isMergeCell[$this->_cellCollection->getCurrentAddress()])) {
                        // Calculated value
                        // To formatted string
						$cellValue = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString(
							$cell->getCalculatedValue(),
							$this->getParent()->getCellXfByIndex($cell->getXfIndex())->getNumberFormat()->getFormatCode()
						);

						$autoSizes[$this->_cellCollection->getCurrentColumn()] = max(
							(float) $autoSizes[$this->_cellCollection->getCurrentColumn()],
                            (float)\PhpOffice\PhpSpreadsheet\Shared\Font::calculateColumnWidth(
								$this->getParent()->getCellXfByIndex($cell->getXfIndex())->getFont(),
                                $cellValue,
								$this->getParent()->getCellXfByIndex($cell->getXfIndex())->getAlignment()->getTextRotation(),
                                $this->getDefaultStyle()->getFont()
                            )
                        );
                    }
                }
            }

            // adjust column widths
            foreach ($autoSizes as $columnIndex => $width) {
                if ($width == -1) $width = $this->getDefaultColumnDimension()->getWidth();
                $this->getColumnDimension($columnIndex)->setWidth($width);
            }
        }

        return $this;
    }

    /**
     * Get parent
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
	public function getParent() {
        return $this->_parent;
    }

    /**
     * Re-bind parent
     *
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $parent
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
	public function rebindParent(\PhpOffice\PhpSpreadsheet\Spreadsheet $parent) {
        if ($this->_parent !== null) {
            $namedRanges = $this->_parent->getNamedRanges();
            foreach ($namedRanges as $namedRange) {
                $parent->addNamedRange($namedRange);
            }

            $this->_parent->removeSheetByIndex(
                $this->_parent->getIndex($this)
            );
        }
        $this->_parent = $parent;

        return $this;
    }

    /**
     * Get title
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->_title;
    }

    /**
     * Set title
     *
     * @param string $pValue String containing the dimension of this worksheet
     * @param string $updateFormulaCellReferences boolean Flag indicating whether cell references in formulae should
     *        	be updated to reflect the new sheet name.
     *          This should be left as the default true, unless you are
     *          certain that no formula cells on any worksheet contain
     *          references to this worksheet
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setTitle($pValue = 'Worksheet', $updateFormulaCellReferences = true)
    {
        // Is this a 'rename' or not?
        if ($this->getTitle() == $pValue) {
            return $this;
        }

        // Syntax check
        self::_checkSheetTitle($pValue);

        // Old title
        $oldTitle = $this->getTitle();

        if ($this->_parent) {
            // Is there already such sheet name?
			if ($this->_parent->sheetNameExists($pValue)) {
                // Use name, but append with lowest possible integer

                if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 29) {
                    $pValue = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,29);
                }
                $i = 1;
				while ($this->_parent->sheetNameExists($pValue . ' ' . $i)) {
                    ++$i;
                    if ($i == 10) {
                        if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 28) {
                            $pValue = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,28);
                        }
                    } elseif ($i == 100) {
                        if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 27) {
                            $pValue = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,27);
                        }
                    }
                }

                $altTitle = $pValue . ' ' . $i;
                return $this->setTitle($altTitle,$updateFormulaCellReferences);
            }
        }

        // Set title
        $this->_title = $pValue;
        $this->_dirty = true;

        if ($this->_parent) {
            // New title
            $newTitle = $this->getTitle();
			\PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($this->_parent)
			    ->renameCalculationCacheForWorksheet($oldTitle, $newTitle);
            if ($updateFormulaCellReferences)
				\PhpOffice\PhpSpreadsheet\ReferenceHelper::getInstance()->updateNamedFormulas($this->_parent, $oldTitle, $newTitle);
        }

        return $this;
    }

    /**
     * Get sheet state
     *
     * @return string Sheet state (visible, hidden, veryHidden)
     */
	public function getSheetState() {
        return $this->_sheetState;
    }

    /**
     * Set sheet state
     *
     * @param string $value Sheet state (visible, hidden, veryHidden)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
	public function setSheetState($value = \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::SHEETSTATE_VISIBLE) {
        $this->_sheetState = $value;
        return $this;
    }

    /**
     * Get page setup
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup
     */
    public function getPageSetup()
    {
        return $this->_pageSetup;
    }

    /**
     * Set page setup
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup    $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setPageSetup(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup $pValue)
    {
        $this->_pageSetup = $pValue;
        return $this;
    }

    /**
     * Get page margins
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\PageMargins
     */
    public function getPageMargins()
    {
        return $this->_pageMargins;
    }

    /**
     * Set page margins
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\PageMargins    $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setPageMargins(\PhpOffice\PhpSpreadsheet\Worksheet\PageMargins $pValue)
    {
        $this->_pageMargins = $pValue;
        return $this;
    }

    /**
     * Get page header/footer
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter
     */
    public function getHeaderFooter()
    {
        return $this->_headerFooter;
    }

    /**
     * Set page header/footer
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter    $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setHeaderFooter(\PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter $pValue)
    {
        $this->_headerFooter = $pValue;
        return $this;
    }

    /**
     * Get sheet view
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\SheetView
     */
    public function getSheetView()
    {
        return $this->_sheetView;
    }

    /**
     * Set sheet view
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\SheetView    $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setSheetView(\PhpOffice\PhpSpreadsheet\Worksheet\SheetView $pValue)
    {
        $this->_sheetView = $pValue;
        return $this;
    }

    /**
     * Get Protection
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function getProtection()
    {
        return $this->_protection;
    }

    /**
     * Set Protection
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Protection    $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setProtection(\PhpOffice\PhpSpreadsheet\Worksheet\Protection $pValue)
    {
        $this->_protection = $pValue;
        $this->_dirty = true;

        return $this;
    }

    /**
     * Get highest worksheet column
     *
     * @param   string     $row        Return the data highest column for the specified row,
     *                                     or the highest column of any row if no row number is passed
     * @return string Highest column name
     */
    public function getHighestColumn($row = null)
    {
        if ($row == null) {
            return $this->_cachedHighestColumn;
        }
        return $this->getHighestDataColumn($row);
    }

    /**
     * Get highest worksheet column that contains data
     *
     * @param   string     $row        Return the highest data column for the specified row,
     *                                     or the highest data column of any row if no row number is passed
     * @return string Highest column name that contains data
     */
    public function getHighestDataColumn($row = null)
    {
        return $this->_cellCollection->getHighestColumn($row);
    }

    /**
     * Get highest worksheet row
     *
     * @param   string     $column     Return the highest data row for the specified column,
     *                                     or the highest row of any column if no column letter is passed
     * @return int Highest row number
     */
    public function getHighestRow($column = null)
    {
        if ($column == null) {
            return $this->_cachedHighestRow;
        }
        return $this->getHighestDataRow($column);
    }

    /**
     * Get highest worksheet row that contains data
     *
     * @param   string     $column     Return the highest data row for the specified column,
     *                                     or the highest data row of any column if no column letter is passed
     * @return string Highest row number that contains data
     */
    public function getHighestDataRow($column = null)
    {
        return $this->_cellCollection->getHighestRow($column);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        return $this->_cellCollection->getHighestRowAndColumn();
    }

    /**
     * Set a cell value
     *
     * @param string $pCoordinate Coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param bool $returnCell   Return the worksheet (false, default) or the cell (true)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|\PhpOffice\PhpSpreadsheet\Cell\Cell    Depending on the last parameter being specified
     */
    public function setCellValue($pCoordinate = 'A1', $pValue = null, $returnCell = false)
    {
        $cell = $this->getCell(strtoupper($pCoordinate))->setValue($pValue);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell (A = 0)
     * @param string $pRow Numeric row coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|\PhpOffice\PhpSpreadsheet\Cell\Cell    Depending on the last parameter being specified
     */
    public function setCellValueByColumnAndRow($pColumn = 0, $pRow = 1, $pValue = null, $returnCell = false)
    {
        $cell = $this->getCellByColumnAndRow($pColumn, $pRow)->setValue($pValue);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Set a cell value
     *
     * @param string $pCoordinate Coordinate of the cell
     * @param mixed  $pValue Value of the cell
     * @param string $pDataType Explicit data type
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|\PhpOffice\PhpSpreadsheet\Cell\Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicit($pCoordinate = 'A1', $pValue = null, $pDataType = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING, $returnCell = false)
    {
        // Set value
        $cell = $this->getCell(strtoupper($pCoordinate))->setValueExplicit($pValue, $pDataType);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param string $pDataType Explicit data type
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|\PhpOffice\PhpSpreadsheet\Cell\Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicitByColumnAndRow($pColumn = 0, $pRow = 1, $pValue = null, $pDataType = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING, $returnCell = false)
    {
        $cell = $this->getCellByColumnAndRow($pColumn, $pRow)->setValueExplicit($pValue, $pDataType);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param string $pCoordinate    Coordinate of the cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Cell\Cell Cell that was found
     */
    public function getCell($pCoordinate = 'A1')
    {
        $pCoordinate = strtoupper($pCoordinate);
        // Check cell collection
        if ($this->_cellCollection->isDataSet($pCoordinate)) {
            return $this->_cellCollection->getCacheData($pCoordinate);
        }

        // Worksheet reference?
        if (strpos($pCoordinate, '!') !== false) {
            $worksheetReference = \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::extractSheetTitle($pCoordinate, true);
			return $this->_parent->getSheetByName($worksheetReference[0])->getCell($worksheetReference[1]);
        }

        // Named range?
        if ((!preg_match('/^'.\PhpOffice\PhpSpreadsheet\Calculation\Calculation::CALCULATION_REGEXP_CELLREF.'$/i', $pCoordinate, $matches)) &&
            (preg_match('/^'.\PhpOffice\PhpSpreadsheet\Calculation\Calculation::CALCULATION_REGEXP_NAMEDRANGE.'$/i', $pCoordinate, $matches))) {
            $namedRange = \PhpOffice\PhpSpreadsheet\NamedRange::resolveRange($pCoordinate, $this);
            if ($namedRange !== NULL) {
                $pCoordinate = $namedRange->getRange();
                return $namedRange->getWorksheet()->getCell($pCoordinate);
            }
        }

        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        if (strpos($pCoordinate, ':') !== false || strpos($pCoordinate, ',') !== false) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate can not be a range of cells.');
        } elseif (strpos($pCoordinate, '$') !== false) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate must not be absolute.');
        }

        // Create new cell object
        return $this->_createNewCell($pCoordinate);
    }

    /**
     * Get cell at a specific coordinate by using numeric cell coordinates
     *
     * @param  string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @return \PhpOffice\PhpSpreadsheet\Cell\Cell Cell that was found
     */
    public function getCellByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        $columnLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn);
        $coordinate = $columnLetter . $pRow;

        if ($this->_cellCollection->isDataSet($coordinate)) {
            return $this->_cellCollection->getCacheData($coordinate);
        }

		return $this->_createNewCell($coordinate);
    }

    /**
     * Create a new cell at the specified coordinate
     *
     * @param string $pCoordinate    Coordinate of the cell
     * @return \PhpOffice\PhpSpreadsheet\Cell\Cell Cell that was created
     */
	private function _createNewCell($pCoordinate)
	{
		$cell = $this->_cellCollection->addCacheData(
			$pCoordinate,
			new \PhpOffice\PhpSpreadsheet\Cell\Cell(
				NULL, 
				\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL, 
				$this
			)
		);
        $this->_cellCollectionIsSorted = false;

        // Coordinates
        $aCoordinates = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($pCoordinate);
        if (\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($this->_cachedHighestColumn) < \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($aCoordinates[0]))
            $this->_cachedHighestColumn = $aCoordinates[0];
        $this->_cachedHighestRow = max($this->_cachedHighestRow, $aCoordinates[1]);

        // Cell needs appropriate xfIndex from dimensions records
		//    but don't create dimension records if they don't already exist
        $rowDimension    = $this->getRowDimension($aCoordinates[1], FALSE);
        $columnDimension = $this->getColumnDimension($aCoordinates[0], FALSE);

        if ($rowDimension !== NULL && $rowDimension->getXfIndex() > 0) {
            // then there is a row dimension with explicit style, assign it to the cell
            $cell->setXfIndex($rowDimension->getXfIndex());
        } elseif ($columnDimension !== NULL && $columnDimension->getXfIndex() > 0) {
            // then there is a column dimension, assign it to the cell
            $cell->setXfIndex($columnDimension->getXfIndex());
        }

        return $cell;
	}
	
    /**
     * Does the cell at a specific coordinate exist?
     *
     * @param string $pCoordinate  Coordinate of the cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return boolean
     */
    public function cellExists($pCoordinate = 'A1')
    {
       // Worksheet reference?
        if (strpos($pCoordinate, '!') !== false) {
            $worksheetReference = \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::extractSheetTitle($pCoordinate, true);
			return $this->_parent->getSheetByName($worksheetReference[0])->cellExists(strtoupper($worksheetReference[1]));
        }

        // Named range?
        if ((!preg_match('/^'.\PhpOffice\PhpSpreadsheet\Calculation\Calculation::CALCULATION_REGEXP_CELLREF.'$/i', $pCoordinate, $matches)) &&
            (preg_match('/^'.\PhpOffice\PhpSpreadsheet\Calculation\Calculation::CALCULATION_REGEXP_NAMEDRANGE.'$/i', $pCoordinate, $matches))) {
            $namedRange = \PhpOffice\PhpSpreadsheet\NamedRange::resolveRange($pCoordinate, $this);
            if ($namedRange !== NULL) {
                $pCoordinate = $namedRange->getRange();
                if ($this->getHashCode() != $namedRange->getWorksheet()->getHashCode()) {
                    if (!$namedRange->getLocalOnly()) {
                        return $namedRange->getWorksheet()->cellExists($pCoordinate);
                    } else {
                        throw new \PhpOffice\PhpSpreadsheet\Exception('Named range ' . $namedRange->getName() . ' is not accessible from within sheet ' . $this->getTitle());
                    }
                }
            }
            else { return false; }
        }

        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        if (strpos($pCoordinate,':') !== false || strpos($pCoordinate,',') !== false) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate can not be a range of cells.');
        } elseif (strpos($pCoordinate,'$') !== false) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate must not be absolute.');
        } else {
            // Coordinates
            $aCoordinates = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($pCoordinate);

            // Cell exists?
            return $this->_cellCollection->isDataSet($pCoordinate);
        }
    }

    /**
     * Cell at a specific coordinate by using numeric cell coordinates exists?
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @return boolean
     */
    public function cellExistsByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->cellExists(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get row dimension at a specific row
     *
     * @param int $pRow Numeric index of the row
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension
     */
    public function getRowDimension($pRow = 1, $create = TRUE)
    {
        // Found
        $found = null;

        // Get row dimension
        if (!isset($this->_rowDimensions[$pRow])) {
			if (!$create)
				return NULL;
            $this->_rowDimensions[$pRow] = new \PhpOffice\PhpSpreadsheet\Worksheet\RowDimension($pRow);

            $this->_cachedHighestRow = max($this->_cachedHighestRow,$pRow);
        }
        return $this->_rowDimensions[$pRow];
    }

    /**
     * Get column dimension at a specific column
     *
     * @param string $pColumn String index of the column
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension
     */
    public function getColumnDimension($pColumn = 'A', $create = TRUE)
    {
        // Uppercase coordinate
        $pColumn = strtoupper($pColumn);

        // Fetch dimensions
        if (!isset($this->_columnDimensions[$pColumn])) {
			if (!$create)
				return NULL;
            $this->_columnDimensions[$pColumn] = new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension($pColumn);

            if (\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($this->_cachedHighestColumn) < \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($pColumn))
                $this->_cachedHighestColumn = $pColumn;
        }
        return $this->_columnDimensions[$pColumn];
    }

    /**
     * Get column dimension at a specific column by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension
     */
    public function getColumnDimensionByColumn($pColumn = 0)
    {
        return $this->getColumnDimension(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn));
    }

    /**
     * Get styles
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style[]
     */
    public function getStyles()
    {
        return $this->_styles;
    }

    /**
     * Get default style of workbook.
     *
     * @deprecated
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getDefaultStyle()
    {
        return $this->_parent->getDefaultStyle();
    }

    /**
     * Set default style - should only be used by \PhpOffice\PhpSpreadsheet\Spreadsheet_IReader implementations!
     *
     * @deprecated
     * @param \PhpOffice\PhpSpreadsheet\Style\Style $pValue
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setDefaultStyle(\PhpOffice\PhpSpreadsheet\Style\Style $pValue)
    {
        $this->_parent->getDefaultStyle()->applyFromArray(array(
            'font' => array(
                'name' => $pValue->getFont()->getName(),
                'size' => $pValue->getFont()->getSize(),
            ),
        ));
        return $this;
    }

    /**
     * Get style for cell
     *
     * @param string $pCellCoordinate Cell coordinate (or range) to get style for
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getStyle($pCellCoordinate = 'A1')
    {
        // set this sheet as active
        $this->_parent->setActiveSheetIndex($this->_parent->getIndex($this));

        // set cell coordinate as active
        $this->setSelectedCells(strtoupper($pCellCoordinate));

        return $this->_parent->getCellXfSupervisor();
    }

    /**
     * Get conditional styles for a cell
     *
     * @param string $pCoordinate
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional[]
     */
    public function getConditionalStyles($pCoordinate = 'A1')
    {
        $pCoordinate = strtoupper($pCoordinate);
        if (!isset($this->_conditionalStylesCollection[$pCoordinate])) {
            $this->_conditionalStylesCollection[$pCoordinate] = array();
        }
        return $this->_conditionalStylesCollection[$pCoordinate];
    }

    /**
     * Do conditional styles exist for this cell?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function conditionalStylesExists($pCoordinate = 'A1')
    {
        if (isset($this->_conditionalStylesCollection[strtoupper($pCoordinate)])) {
            return true;
        }
        return false;
    }

    /**
     * Removes conditional styles for a cell
     *
     * @param string $pCoordinate
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function removeConditionalStyles($pCoordinate = 'A1')
    {
        unset($this->_conditionalStylesCollection[strtoupper($pCoordinate)]);
        return $this;
    }

    /**
     * Get collection of conditional styles
     *
     * @return array
     */
    public function getConditionalStylesCollection()
    {
        return $this->_conditionalStylesCollection;
    }

    /**
     * Set conditional styles
     *
     * @param $pCoordinate string E.g. 'A1'
     * @param $pValue \PhpOffice\PhpSpreadsheet\Style\Conditional[]
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setConditionalStyles($pCoordinate = 'A1', $pValue)
    {
        $this->_conditionalStylesCollection[strtoupper($pCoordinate)] = $pValue;
        return $this;
    }

    /**
     * Get style for cell by using numeric cell coordinates
     *
     * @param int $pColumn  Numeric column coordinate of the cell
     * @param int $pRow Numeric row coordinate of the cell
     * @param int pColumn2 Numeric column coordinate of the range cell
     * @param int pRow2 Numeric row coordinate of the range cell
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function getStyleByColumnAndRow($pColumn = 0, $pRow = 1, $pColumn2 = null, $pRow2 = null)
    {
        if (!is_null($pColumn2) && !is_null($pRow2)) {
		    $cellRange = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow . ':' . 
                \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn2) . $pRow2;
		    return $this->getStyle($cellRange);
	    }

        return $this->getStyle(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Set shared cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @deprecated
     * @param \PhpOffice\PhpSpreadsheet\Style\Style $pSharedCellStyle Cell style to share
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setSharedStyle(\PhpOffice\PhpSpreadsheet\Style\Style $pSharedCellStyle = null, $pRange = '')
    {
        $this->duplicateStyle($pSharedCellStyle, $pRange);
        return $this;
    }

    /**
     * Duplicate cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param \PhpOffice\PhpSpreadsheet\Style\Style $pCellStyle Cell style to duplicate
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function duplicateStyle(\PhpOffice\PhpSpreadsheet\Style\Style $pCellStyle = null, $pRange = '')
    {
        // make sure we have a real style and not supervisor
        $style = $pCellStyle->getIsSupervisor() ? $pCellStyle->getSharedComponent() : $pCellStyle;

        // Add the style to the workbook if necessary
        $workbook = $this->_parent;
		if ($existingStyle = $this->_parent->getCellXfByHashCode($pCellStyle->getHashCode())) {
            // there is already such cell Xf in our collection
            $xfIndex = $existingStyle->getIndex();
        } else {
            // we don't have such a cell Xf, need to add
            $workbook->addCellXf($pCellStyle);
            $xfIndex = $pCellStyle->getIndex();
        }

        // Calculate range outer borders
        list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($pRange . ':' . $pRange);

        // Make sure we can loop upwards on rows and columns
        if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
            $tmp = $rangeStart;
            $rangeStart = $rangeEnd;
            $rangeEnd = $tmp;
        }

        // Loop through cells and apply styles
        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->getCell(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col - 1) . $row)->setXfIndex($xfIndex);
            }
        }

        return $this;
    }

    /**
     * Duplicate conditional style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
	 * @param	array of \PhpOffice\PhpSpreadsheet\Style\Conditional	$pCellStyle	Cell style to duplicate
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function duplicateConditionalStyle(array $pCellStyle = null, $pRange = '')
    {
        foreach($pCellStyle as $cellStyle) {
            if (!($cellStyle instanceof \PhpOffice\PhpSpreadsheet\Style\Conditional)) {
                throw new \PhpOffice\PhpSpreadsheet\Exception('Style is not a conditional style');
            }
        }

        // Calculate range outer borders
        list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($pRange . ':' . $pRange);

        // Make sure we can loop upwards on rows and columns
        if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
            $tmp = $rangeStart;
            $rangeStart = $rangeEnd;
            $rangeEnd = $tmp;
        }

        // Loop through cells and apply styles
        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->setConditionalStyles(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col - 1) . $row, $pCellStyle);
            }
        }

        return $this;
    }

    /**
     * Duplicate cell style array to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range,
     * if they are in the styles array. For example, if you decide to set a range of
     * cells to font bold, only include font bold in the styles array.
     *
     * @deprecated
     * @param array $pStyles Array containing style information
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param boolean $pAdvanced Advanced mode for setting borders.
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function duplicateStyleArray($pStyles = null, $pRange = '', $pAdvanced = true)
    {
        $this->getStyle($pRange)->applyFromArray($pStyles, $pAdvanced);
        return $this;
    }

    /**
     * Set break on a cell
     *
     * @param string $pCell Cell coordinate (e.g. A1)
     * @param int $pBreak Break type (type of \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_*)
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setBreak($pCell = 'A1', $pBreak = \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_NONE)
    {
        // Uppercase coordinate
        $pCell = strtoupper($pCell);

        if ($pCell != '') {
        	if ($pBreak == \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_NONE) {
        		if (isset($this->_breaks[$pCell])) {
	            	unset($this->_breaks[$pCell]);
        		}
        	} else {
	            $this->_breaks[$pCell] = $pBreak;
	        }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception('No cell coordinate specified.');
        }

        return $this;
    }

    /**
     * Set break on a cell by using numeric cell coordinates
     *
     * @param integer $pColumn Numeric column coordinate of the cell
     * @param integer $pRow Numeric row coordinate of the cell
     * @param  integer $pBreak Break type (type of \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_*)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setBreakByColumnAndRow($pColumn = 0, $pRow = 1, $pBreak = \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_NONE)
    {
        return $this->setBreak(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow, $pBreak);
    }

    /**
     * Get breaks
     *
     * @return array[]
     */
    public function getBreaks()
    {
        return $this->_breaks;
    }

    /**
     * Set merge on a cell range
     *
     * @param string $pRange  Cell range (e.g. A1:E1)
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function mergeCells($pRange = 'A1:A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (strpos($pRange,':') !== false) {
            $this->_mergeCells[$pRange] = $pRange;

            // make sure cells are created

            // get the cells in the range
            $aReferences = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($pRange);

            // create upper left cell if it does not already exist
            $upperLeft = $aReferences[0];
            if (!$this->cellExists($upperLeft)) {
                $this->getCell($upperLeft)->setValueExplicit(null, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL);
            }

            // create or blank out the rest of the cells in the range
            $count = count($aReferences);
            for ($i = 1; $i < $count; $i++) {
                $this->getCell($aReferences[$i])->setValueExplicit(null, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL);
            }

        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Merge must be set on a range of cells.');
        }

        return $this;
    }

    /**
     * Set merge on a cell range by using numeric cell coordinates
     *
     * @param int $pColumn1    Numeric column coordinate of the first cell
     * @param int $pRow1        Numeric row coordinate of the first cell
     * @param int $pColumn2    Numeric column coordinate of the last cell
     * @param int $pRow2        Numeric row coordinate of the last cell
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function mergeCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $cellRange = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->mergeCells($cellRange);
    }

    /**
     * Remove merge on a cell range
     *
     * @param    string            $pRange        Cell range (e.g. A1:E1)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function unmergeCells($pRange = 'A1:A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (strpos($pRange,':') !== false) {
            if (isset($this->_mergeCells[$pRange])) {
                unset($this->_mergeCells[$pRange]);
            } else {
                throw new \PhpOffice\PhpSpreadsheet\Exception('Cell range ' . $pRange . ' not known as merged.');
            }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Merge can only be removed from a range of cells.');
        }

        return $this;
    }

    /**
     * Remove merge on a cell range by using numeric cell coordinates
     *
     * @param int $pColumn1    Numeric column coordinate of the first cell
     * @param int $pRow1        Numeric row coordinate of the first cell
     * @param int $pColumn2    Numeric column coordinate of the last cell
     * @param int $pRow2        Numeric row coordinate of the last cell
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function unmergeCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $cellRange = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->unmergeCells($cellRange);
    }

    /**
     * Get merge cells array.
     *
     * @return array[]
     */
    public function getMergeCells()
    {
        return $this->_mergeCells;
    }

    /**
     * Set merge cells array for the entire sheet. Use instead mergeCells() to merge
     * a single cell range.
     *
     * @param array
     */
    public function setMergeCells($pValue = array())
    {
        $this->_mergeCells = $pValue;

        return $this;
    }

    /**
     * Set protection on a cell range
     *
     * @param    string            $pRange                Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @param    string            $pPassword            Password to unlock the protection
     * @param    boolean        $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function protectCells($pRange = 'A1', $pPassword = '', $pAlreadyHashed = false)
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (!$pAlreadyHashed) {
            $pPassword = \PhpOffice\PhpSpreadsheet\Shared\PasswordHasher::hashPassword($pPassword);
        }
        $this->_protectedCells[$pRange] = $pPassword;

        return $this;
    }

    /**
     * Set protection on a cell range by using numeric cell coordinates
     *
     * @param int  $pColumn1            Numeric column coordinate of the first cell
     * @param int  $pRow1                Numeric row coordinate of the first cell
     * @param int  $pColumn2            Numeric column coordinate of the last cell
     * @param int  $pRow2                Numeric row coordinate of the last cell
     * @param string $pPassword            Password to unlock the protection
     * @param    boolean $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function protectCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1, $pPassword = '', $pAlreadyHashed = false)
    {
        $cellRange = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->protectCells($cellRange, $pPassword, $pAlreadyHashed);
    }

    /**
     * Remove protection on a cell range
     *
     * @param    string            $pRange        Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function unprotectCells($pRange = 'A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (isset($this->_protectedCells[$pRange])) {
            unset($this->_protectedCells[$pRange]);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell range ' . $pRange . ' not known as protected.');
        }
        return $this;
    }

    /**
     * Remove protection on a cell range by using numeric cell coordinates
     *
     * @param int  $pColumn1            Numeric column coordinate of the first cell
     * @param int  $pRow1                Numeric row coordinate of the first cell
     * @param int  $pColumn2            Numeric column coordinate of the last cell
     * @param int $pRow2                Numeric row coordinate of the last cell
     * @param string $pPassword            Password to unlock the protection
     * @param    boolean $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function unprotectCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1, $pPassword = '', $pAlreadyHashed = false)
    {
        $cellRange = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->unprotectCells($cellRange, $pPassword, $pAlreadyHashed);
    }

    /**
     * Get protected cells
     *
     * @return array[]
     */
    public function getProtectedCells()
    {
        return $this->_protectedCells;
    }

    /**
     *    Get Autofilter
     *
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function getAutoFilter()
    {
        return $this->_autoFilter;
    }

    /**
     *    Set AutoFilter
     *
     *    @param    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter|string   $pValue
     *            A simple string containing a Cell range like 'A1:E10' is permitted for backward compatibility
     *    @throws    \PhpOffice\PhpSpreadsheet\Exception
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setAutoFilter($pValue)
    {
        $pRange = strtoupper($pValue);

        if (is_string($pValue)) {
            $this->_autoFilter->setRange($pValue);
        } elseif(is_object($pValue) && ($pValue instanceof \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter)) {
            $this->_autoFilter = $pValue;
        }
        return $this;
    }

    /**
     *    Set Autofilter Range by using numeric cell coordinates
     *
     *    @param  integer  $pColumn1    Numeric column coordinate of the first cell
     *    @param  integer  $pRow1       Numeric row coordinate of the first cell
     *    @param  integer  $pColumn2    Numeric column coordinate of the second cell
     *    @param  integer  $pRow2       Numeric row coordinate of the second cell
     *    @throws    \PhpOffice\PhpSpreadsheet\Exception
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setAutoFilterByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        return $this->setAutoFilter(
            \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn1) . $pRow1
            . ':' .
            \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn2) . $pRow2
        );
    }

    /**
     * Remove autofilter
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function removeAutoFilter()
    {
        $this->_autoFilter->setRange(NULL);
        return $this;
    }

    /**
     * Get Freeze Pane
     *
     * @return string
     */
    public function getFreezePane()
    {
        return $this->_freezePane;
    }

    /**
     * Freeze Pane
     *
     * @param    string        $pCell        Cell (i.e. A2)
     *                                    Examples:
     *                                        A2 will freeze the rows above cell A2 (i.e row 1)
     *                                        B1 will freeze the columns to the left of cell B1 (i.e column A)
     *                                        B2 will freeze the rows above and to the left of cell A2
     *                                            (i.e row 1 and column A)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function freezePane($pCell = '')
    {
        // Uppercase coordinate
        $pCell = strtoupper($pCell);

        if (strpos($pCell,':') === false && strpos($pCell,',') === false) {
            $this->_freezePane = $pCell;
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Freeze pane can not be set on a range of cells.');
        }
        return $this;
    }

    /**
     * Freeze Pane by using numeric cell coordinates
     *
     * @param int $pColumn    Numeric column coordinate of the cell
     * @param int $pRow        Numeric row coordinate of the cell
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function freezePaneByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->freezePane(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Unfreeze Pane
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function unfreezePane()
    {
        return $this->freezePane('');
    }

    /**
     * Insert a new row, updating all possible related data
     *
     * @param int $pBefore    Insert before this one
     * @param int $pNumRows    Number of rows to insert
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function insertNewRowBefore($pBefore = 1, $pNumRows = 1) {
        if ($pBefore >= 1) {
            $objReferenceHelper = \PhpOffice\PhpSpreadsheet\ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore('A' . $pBefore, 0, $pNumRows, $this);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Rows can only be inserted before at least row 1.");
        }
        return $this;
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int $pBefore    Insert before this one
     * @param int $pNumCols    Number of columns to insert
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function insertNewColumnBefore($pBefore = 'A', $pNumCols = 1) {
        if (!is_numeric($pBefore)) {
            $objReferenceHelper = \PhpOffice\PhpSpreadsheet\ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore($pBefore . '1', $pNumCols, 0, $this);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Column references should not be numeric.");
        }
        return $this;
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int $pBefore    Insert before this one (numeric column coordinate of the cell)
     * @param int $pNumCols    Number of columns to insert
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function insertNewColumnBeforeByIndex($pBefore = 0, $pNumCols = 1) {
        if ($pBefore >= 0) {
            return $this->insertNewColumnBefore(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pBefore), $pNumCols);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Columns can only be inserted before at least column A (0).");
        }
    }

    /**
     * Delete a row, updating all possible related data
     *
     * @param int $pRow        Remove starting with this one
     * @param int $pNumRows    Number of rows to remove
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function removeRow($pRow = 1, $pNumRows = 1) {
        if ($pRow >= 1) {
            $highestRow = $this->getHighestDataRow();
            $objReferenceHelper = \PhpOffice\PhpSpreadsheet\ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore('A' . ($pRow + $pNumRows), 0, -$pNumRows, $this);
            for($r = 0; $r < $pNumRows; ++$r) {
                $this->getCellCacheController()->removeRow($highestRow);
                --$highestRow;
            }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Rows to be deleted should at least start from row 1.");
        }
        return $this;
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param string    $pColumn     Remove starting with this one
     * @param int       $pNumCols    Number of columns to remove
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function removeColumn($pColumn = 'A', $pNumCols = 1) {
        if (!is_numeric($pColumn)) {
            $highestColumn = $this->getHighestDataColumn();
            $pColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($pColumn) - 1 + $pNumCols);
            $objReferenceHelper = \PhpOffice\PhpSpreadsheet\ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore($pColumn . '1', -$pNumCols, 0, $this);
            for($c = 0; $c < $pNumCols; ++$c) {
                $this->getCellCacheController()->removeColumn($highestColumn);
                $highestColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn) - 2);
            }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Column references should not be numeric.");
        }
        return $this;
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param int $pColumn    Remove starting with this one (numeric column coordinate of the cell)
     * @param int $pNumCols    Number of columns to remove
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function removeColumnByIndex($pColumn = 0, $pNumCols = 1) {
        if ($pColumn >= 0) {
            return $this->removeColumn(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn), $pNumCols);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Columns to be deleted should at least start from column 0");
        }
    }

    /**
     * Show gridlines?
     *
     * @return boolean
     */
    public function getShowGridlines() {
        return $this->_showGridlines;
    }

    /**
     * Set show gridlines
     *
     * @param boolean $pValue    Show gridlines (true/false)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setShowGridlines($pValue = false) {
        $this->_showGridlines = $pValue;
        return $this;
    }

    /**
    * Print gridlines?
    *
    * @return boolean
    */
    public function getPrintGridlines() {
        return $this->_printGridlines;
    }

    /**
    * Set print gridlines
    *
    * @param boolean $pValue Print gridlines (true/false)
    * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    */
    public function setPrintGridlines($pValue = false) {
        $this->_printGridlines = $pValue;
        return $this;
    }

    /**
    * Show row and column headers?
    *
    * @return boolean
    */
    public function getShowRowColHeaders() {
        return $this->_showRowColHeaders;
    }

    /**
    * Set show row and column headers
    *
    * @param boolean $pValue Show row and column headers (true/false)
    * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    */
    public function setShowRowColHeaders($pValue = false) {
        $this->_showRowColHeaders = $pValue;
        return $this;
    }

    /**
     * Show summary below? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryBelow() {
        return $this->_showSummaryBelow;
    }

    /**
     * Set show summary below
     *
     * @param boolean $pValue    Show summary below (true/false)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setShowSummaryBelow($pValue = true) {
        $this->_showSummaryBelow = $pValue;
        return $this;
    }

    /**
     * Show summary right? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryRight() {
        return $this->_showSummaryRight;
    }

    /**
     * Set show summary right
     *
     * @param boolean $pValue    Show summary right (true/false)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setShowSummaryRight($pValue = true) {
        $this->_showSummaryRight = $pValue;
        return $this;
    }

    /**
     * Get comments
     *
     * @return \PhpOffice\PhpSpreadsheet\Comment[]
     */
    public function getComments()
    {
        return $this->_comments;
    }

    /**
     * Set comments array for the entire sheet.
     *
	 * @param array of \PhpOffice\PhpSpreadsheet\Comment
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setComments($pValue = array())
    {
        $this->_comments = $pValue;

        return $this;
    }

    /**
     * Get comment for cell
     *
     * @param string $pCellCoordinate    Cell coordinate to get comment for
     * @return \PhpOffice\PhpSpreadsheet\Comment
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getComment($pCellCoordinate = 'A1')
    {
        // Uppercase coordinate
        $pCellCoordinate = strtoupper($pCellCoordinate);

        if (strpos($pCellCoordinate,':') !== false || strpos($pCellCoordinate,',') !== false) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate string can not be a range of cells.');
        } else if (strpos($pCellCoordinate,'$') !== false) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate string must not be absolute.');
        } else if ($pCellCoordinate == '') {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Cell coordinate can not be zero-length string.');
        } else {
            // Check if we already have a comment for this cell.
            // If not, create a new comment.
            if (isset($this->_comments[$pCellCoordinate])) {
                return $this->_comments[$pCellCoordinate];
            } else {
                $newComment = new \PhpOffice\PhpSpreadsheet\Comment();
                $this->_comments[$pCellCoordinate] = $newComment;
                return $newComment;
            }
        }
    }

    /**
     * Get comment for cell by using numeric cell coordinates
     *
     * @param int $pColumn    Numeric column coordinate of the cell
     * @param int $pRow        Numeric row coordinate of the cell
     * @return \PhpOffice\PhpSpreadsheet\Comment
     */
    public function getCommentByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->getComment(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get selected cell
     *
     * @deprecated
     * @return string
     */
    public function getSelectedCell()
    {
        return $this->getSelectedCells();
    }

    /**
     * Get active cell
     *
     * @return string Example: 'A1'
     */
    public function getActiveCell()
    {
        return $this->_activeCell;
    }

    /**
     * Get selected cells
     *
     * @return string
     */
    public function getSelectedCells()
    {
        return $this->_selectedCells;
    }

    /**
     * Selected cell
     *
     * @param    string        $pCoordinate    Cell (i.e. A1)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setSelectedCell($pCoordinate = 'A1')
    {
        return $this->setSelectedCells($pCoordinate);
    }

    /**
     * Select a range of cells.
     *
     * @param    string        $pCoordinate    Cell range, examples: 'A1', 'B2:G5', 'A:C', '3:6'
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setSelectedCells($pCoordinate = 'A1')
    {
        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        // Convert 'A' to 'A:A'
        $pCoordinate = preg_replace('/^([A-Z]+)$/', '${1}:${1}', $pCoordinate);

        // Convert '1' to '1:1'
        $pCoordinate = preg_replace('/^([0-9]+)$/', '${1}:${1}', $pCoordinate);

        // Convert 'A:C' to 'A1:C1048576'
        $pCoordinate = preg_replace('/^([A-Z]+):([A-Z]+)$/', '${1}1:${2}1048576', $pCoordinate);

        // Convert '1:3' to 'A1:XFD3'
        $pCoordinate = preg_replace('/^([0-9]+):([0-9]+)$/', 'A${1}:XFD${2}', $pCoordinate);

        if (strpos($pCoordinate,':') !== false || strpos($pCoordinate,',') !== false) {
            list($first, ) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($pCoordinate);
            $this->_activeCell = $first[0];
        } else {
            $this->_activeCell = $pCoordinate;
        }
        $this->_selectedCells = $pCoordinate;
        return $this;
    }

    /**
     * Selected cell by using numeric cell coordinates
     *
     * @param int $pColumn Numeric column coordinate of the cell
     * @param int $pRow Numeric row coordinate of the cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setSelectedCellByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->setSelectedCells(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get right-to-left
     *
     * @return boolean
     */
    public function getRightToLeft() {
        return $this->_rightToLeft;
    }

    /**
     * Set right-to-left
     *
     * @param boolean $value    Right-to-left true/false
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setRightToLeft($value = false) {
        $this->_rightToLeft = $value;
        return $this;
    }

    /**
     * Fill worksheet from values in array
     *
     * @param array $source Source array
     * @param mixed $nullValue Value in source array that stands for blank cell
     * @param string $startCell Insert array starting from this cell address as the top left coordinate
     * @param boolean $strictNullComparison Apply strict comparison when testing for null values in the array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function fromArray($source = null, $nullValue = null, $startCell = 'A1', $strictNullComparison = false) {
        if (is_array($source)) {
            //    Convert a 1-D array to 2-D (for ease of looping)
            if (!is_array(end($source))) {
                $source = array($source);
            }

            // start coordinate
            list ($startColumn, $startRow) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($startCell);

            // Loop through $source
            foreach ($source as $rowData) {
                $currentColumn = $startColumn;
                foreach($rowData as $cellValue) {
                    if ($strictNullComparison) {
                        if ($cellValue !== $nullValue) {
                            // Set cell value
                            $this->getCell($currentColumn . $startRow)->setValue($cellValue);
                        }
                    } else {
                        if ($cellValue != $nullValue) {
                            // Set cell value
                            $this->getCell($currentColumn . $startRow)->setValue($cellValue);
                        }
                    }
                    ++$currentColumn;
                }
                ++$startRow;
            }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Parameter \$source should be an array.");
        }
        return $this;
    }

    /**
     * Create array from a range of cells
     *
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
	public function rangeToArray($pRange = 'A1', $nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        // Returnvalue
        $returnValue = array();
        //    Identify the range that we need to extract from the worksheet
        list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($pRange);
        $minCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeStart[0] -1);
        $minRow = $rangeStart[1];
        $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeEnd[0] -1);
        $maxRow = $rangeEnd[1];

        $maxCol++;
        // Loop through rows
        $r = -1;
        for ($row = $minRow; $row <= $maxRow; ++$row) {
            $rRef = ($returnCellRef) ? $row : ++$r;
            $c = -1;
            // Loop through columns in the current row
            for ($col = $minCol; $col != $maxCol; ++$col) {
                $cRef = ($returnCellRef) ? $col : ++$c;
                //    Using getCell() will create a new cell if it doesn't already exist. We don't want that to happen
                //        so we test and retrieve directly against _cellCollection
                if ($this->_cellCollection->isDataSet($col.$row)) {
                    // Cell exists
                    $cell = $this->_cellCollection->getCacheData($col.$row);
                    if ($cell->getValue() !== null) {
                        if ($cell->getValue() instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                            $returnValue[$rRef][$cRef] = $cell->getValue()->getPlainText();
                        } else {
                            if ($calculateFormulas) {
                                $returnValue[$rRef][$cRef] = $cell->getCalculatedValue();
                            } else {
                                $returnValue[$rRef][$cRef] = $cell->getValue();
                            }
                        }

                        if ($formatData) {
                            $style = $this->_parent->getCellXfByIndex($cell->getXfIndex());
                            $returnValue[$rRef][$cRef] = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString(
                            	$returnValue[$rRef][$cRef],
								($style && $style->getNumberFormat()) ?
									$style->getNumberFormat()->getFormatCode() :
									\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_GENERAL
                            );
                        }
                    } else {
                        // Cell holds a NULL
                        $returnValue[$rRef][$cRef] = $nullValue;
                    }
                } else {
                    // Cell doesn't exist
                    $returnValue[$rRef][$cRef] = $nullValue;
                }
            }
        }

        // Return
        return $returnValue;
    }


    /**
     * Create array from a range of cells
     *
     * @param  string $pNamedRange Name of the Named Range
     * @param  mixed  $nullValue Value returned in the array entry if a cell doesn't exist
     * @param  boolean $calculateFormulas  Should formulas be calculated?
     * @param  boolean $formatData  Should formatting be applied to cell values?
     * @param  boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                                True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
	public function namedRangeToArray($pNamedRange = '', $nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        $namedRange = \PhpOffice\PhpSpreadsheet\NamedRange::resolveRange($pNamedRange, $this);
        if ($namedRange !== NULL) {
            $pWorkSheet = $namedRange->getWorksheet();
            $pCellRange = $namedRange->getRange();

			return $pWorkSheet->rangeToArray(	$pCellRange,
												$nullValue, $calculateFormulas, $formatData, $returnCellRef);
        }

        throw new \PhpOffice\PhpSpreadsheet\Exception('Named Range '.$pNamedRange.' does not exist.');
    }


    /**
     * Create array from worksheet
     *
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData  Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
	public function toArray($nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        // Garbage collect...
        $this->garbageCollect();

        //    Identify the range that we need to extract from the worksheet
        $maxCol = $this->getHighestColumn();
        $maxRow = $this->getHighestRow();
        // Return
		return $this->rangeToArray(	'A1:'.$maxCol.$maxRow,
									$nullValue, $calculateFormulas, $formatData, $returnCellRef);
    }

    /**
     * Get row iterator
     *
     * @param   integer   $startRow   The row number at which to start iterating
     * @param   integer   $endRow     The row number at which to stop iterating
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator
     */
	public function getRowIterator($startRow = 1, $endRow = null) {
        return new \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator($this, $startRow, $endRow);
    }

    /**
     * Get column iterator
     *
     * @param   string   $startColumn The column address at which to start iterating
     * @param   string   $endColumn   The column address at which to stop iterating
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator
     */
	public function getColumnIterator($startColumn = 'A', $endColumn = null) {
        return new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnIterator($this, $startColumn, $endColumn);
    }

    /**
     * Run \PhpOffice\PhpSpreadsheet\Spreadsheet garabage collector.
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
	public function garbageCollect() {
        // Flush cache
        $this->_cellCollection->getCacheData('A1');
        // Build a reference table from images
//        $imageCoordinates = array();
//        $iterator = $this->getDrawingCollection()->getIterator();
//        while ($iterator->valid()) {
//            $imageCoordinates[$iterator->current()->getCoordinates()] = true;
//
//            $iterator->next();
//        }
//
        // Lookup highest column and highest row if cells are cleaned
        $colRow = $this->_cellCollection->getHighestRowAndColumn();
        $highestRow = $colRow['row'];
        $highestColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($colRow['column']);

        // Loop through column dimensions
        foreach ($this->_columnDimensions as $dimension) {
            $highestColumn = max($highestColumn,\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($dimension->getColumnIndex()));
        }

        // Loop through row dimensions
        foreach ($this->_rowDimensions as $dimension) {
            $highestRow = max($highestRow,$dimension->getRowIndex());
        }

        // Cache values
        if ($highestColumn < 0) {
            $this->_cachedHighestColumn = 'A';
        } else {
            $this->_cachedHighestColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(--$highestColumn);
        }
        $this->_cachedHighestRow = $highestRow;

        // Return
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
	public function getHashCode() {
        if ($this->_dirty) {
            $this->_hash = md5( $this->_title .
                                $this->_autoFilter .
                                ($this->_protection->isProtectionEnabled() ? 't' : 'f') .
                                __CLASS__
                              );
            $this->_dirty = false;
        }
        return $this->_hash;
    }

    /**
     * Extract worksheet title from range.
     *
     * Example: extractSheetTitle("testSheet!A1") ==> 'A1'
     * Example: extractSheetTitle("'testSheet 1'!A1", true) ==> array('testSheet 1', 'A1');
     *
     * @param string $pRange    Range to extract title from
     * @param bool $returnRange    Return range? (see example)
     * @return mixed
     */
	public static function extractSheetTitle($pRange, $returnRange = false) {
        // Sheet title included?
        if (($sep = strpos($pRange, '!')) === false) {
            return '';
        }

        if ($returnRange) {
            return array( trim(substr($pRange, 0, $sep),"'"),
                          substr($pRange, $sep + 1)
                        );
        }

        return substr($pRange, $sep + 1);
    }

    /**
     * Get hyperlink
     *
     * @param string $pCellCoordinate    Cell coordinate to get hyperlink for
     */
    public function getHyperlink($pCellCoordinate = 'A1')
    {
        // return hyperlink if we already have one
        if (isset($this->_hyperlinkCollection[$pCellCoordinate])) {
            return $this->_hyperlinkCollection[$pCellCoordinate];
        }

        // else create hyperlink
        $this->_hyperlinkCollection[$pCellCoordinate] = new \PhpOffice\PhpSpreadsheet\Cell\Hyperlink();
        return $this->_hyperlinkCollection[$pCellCoordinate];
    }

    /**
     * Set hyperlnk
     *
     * @param string $pCellCoordinate    Cell coordinate to insert hyperlink
     * @param    \PhpOffice\PhpSpreadsheet\Cell\Hyperlink    $pHyperlink
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setHyperlink($pCellCoordinate = 'A1', \PhpOffice\PhpSpreadsheet\Cell\Hyperlink $pHyperlink = null)
    {
        if ($pHyperlink === null) {
            unset($this->_hyperlinkCollection[$pCellCoordinate]);
        } else {
            $this->_hyperlinkCollection[$pCellCoordinate] = $pHyperlink;
        }
        return $this;
    }

    /**
     * Hyperlink at a specific coordinate exists?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function hyperlinkExists($pCoordinate = 'A1')
    {
        return isset($this->_hyperlinkCollection[$pCoordinate]);
    }

    /**
     * Get collection of hyperlinks
     *
     * @return \PhpOffice\PhpSpreadsheet\Cell\Hyperlink[]
     */
    public function getHyperlinkCollection()
    {
        return $this->_hyperlinkCollection;
    }

    /**
     * Get data validation
     *
     * @param string $pCellCoordinate Cell coordinate to get data validation for
     */
    public function getDataValidation($pCellCoordinate = 'A1')
    {
        // return data validation if we already have one
        if (isset($this->_dataValidationCollection[$pCellCoordinate])) {
            return $this->_dataValidationCollection[$pCellCoordinate];
        }

        // else create data validation
        $this->_dataValidationCollection[$pCellCoordinate] = new \PhpOffice\PhpSpreadsheet\Cell\DataValidation();
        return $this->_dataValidationCollection[$pCellCoordinate];
    }

    /**
     * Set data validation
     *
     * @param string $pCellCoordinate    Cell coordinate to insert data validation
     * @param    \PhpOffice\PhpSpreadsheet\Cell\DataValidation    $pDataValidation
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setDataValidation($pCellCoordinate = 'A1', \PhpOffice\PhpSpreadsheet\Cell\DataValidation $pDataValidation = null)
    {
        if ($pDataValidation === null) {
            unset($this->_dataValidationCollection[$pCellCoordinate]);
        } else {
            $this->_dataValidationCollection[$pCellCoordinate] = $pDataValidation;
        }
        return $this;
    }

    /**
     * Data validation at a specific coordinate exists?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function dataValidationExists($pCoordinate = 'A1')
    {
        return isset($this->_dataValidationCollection[$pCoordinate]);
    }

    /**
     * Get collection of data validations
     *
     * @return \PhpOffice\PhpSpreadsheet\Cell\DataValidation[]
     */
    public function getDataValidationCollection()
    {
        return $this->_dataValidationCollection;
    }

    /**
     * Accepts a range, returning it as a range that falls within the current highest row and column of the worksheet
     *
     * @param string $range
     * @return string Adjusted range value
     */
	public function shrinkRangeToFit($range) {
        $maxCol = $this->getHighestColumn();
        $maxRow = $this->getHighestRow();
        $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($maxCol);

        $rangeBlocks = explode(' ',$range);
        foreach ($rangeBlocks as &$rangeSet) {
            $rangeBoundaries = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::getRangeBoundaries($rangeSet);

            if (\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($rangeBoundaries[0][0]) > $maxCol) { $rangeBoundaries[0][0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($maxCol); }
            if ($rangeBoundaries[0][1] > $maxRow) { $rangeBoundaries[0][1] = $maxRow; }
            if (\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($rangeBoundaries[1][0]) > $maxCol) { $rangeBoundaries[1][0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($maxCol); }
            if ($rangeBoundaries[1][1] > $maxRow) { $rangeBoundaries[1][1] = $maxRow; }
            $rangeSet = $rangeBoundaries[0][0].$rangeBoundaries[0][1].':'.$rangeBoundaries[1][0].$rangeBoundaries[1][1];
        }
        unset($rangeSet);
        $stRange = implode(' ',$rangeBlocks);

        return $stRange;
    }

    /**
     * Get tab color
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Color
     */
    public function getTabColor()
    {
        if ($this->_tabColor === NULL)
            $this->_tabColor = new \PhpOffice\PhpSpreadsheet\Style\Color();

        return $this->_tabColor;
    }

    /**
     * Reset tab color
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function resetTabColor()
    {
        $this->_tabColor = null;
        unset($this->_tabColor);

        return $this;
    }

    /**
     * Tab color set?
     *
     * @return boolean
     */
    public function isTabColorSet()
    {
        return ($this->_tabColor !== NULL);
    }

    /**
     * Copy worksheet (!= clone!)
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
	public function copy() {
        $copied = clone $this;

        return $copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
	public function __clone() {
        foreach ($this as $key => $val) {
            if ($key == '_parent') {
                continue;
            }

            if (is_object($val) || (is_array($val))) {
                if ($key == '_cellCollection') {
                    $newCollection = clone $this->_cellCollection;
                    $newCollection->copyCellCollection($this);
                    $this->_cellCollection = $newCollection;
                } elseif ($key == '_drawingCollection') {
                    $newCollection = clone $this->_drawingCollection;
                    $this->_drawingCollection = $newCollection;
                } elseif (($key == '_autoFilter') && ($this->_autoFilter instanceof \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter)) {
                    $newAutoFilter = clone $this->_autoFilter;
                    $this->_autoFilter = $newAutoFilter;
                    $this->_autoFilter->setParent($this);
                } else {
                    $this->{$key} = unserialize(serialize($val));
                }
            }
        }
    }
/**
	 * Define the code name of the sheet
	 *
	 * @param null|string Same rule as Title minus space not allowed (but, like Excel, change silently space to underscore)
	 * @return objWorksheet
	 * @throws \PhpOffice\PhpSpreadsheet\Exception
	*/
	public function setCodeName($pValue=null){
		// Is this a 'rename' or not?
		if ($this->getCodeName() == $pValue) {
			return $this;
		}
		$pValue = str_replace(' ', '_', $pValue);//Excel does this automatically without flinching, we are doing the same
		// Syntax check
        // throw an exception if not valid
		self::_checkSheetCodeName($pValue);

		// We use the same code that setTitle to find a valid codeName else not using a space (Excel don't like) but a '_'
		
        if ($this->getParent()) {
			// Is there already such sheet name?
			if ($this->getParent()->sheetCodeNameExists($pValue)) {
				// Use name, but append with lowest possible integer

				if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 29) {
					$pValue = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,29);
				}
				$i = 1;
				while ($this->getParent()->sheetCodeNameExists($pValue . '_' . $i)) {
					++$i;
					if ($i == 10) {
						if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 28) {
							$pValue = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,28);
						}
					} elseif ($i == 100) {
						if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($pValue) > 27) {
							$pValue = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::Substring($pValue,0,27);
						}
					}
				}

				$pValue = $pValue . '_' . $i;// ok, we have a valid name
				//codeName is'nt used in formula : no need to call for an update
				//return $this->setTitle($altTitle,$updateFormulaCellReferences);
			}
		}

		$this->_codeName=$pValue;
		return $this;
	}
	/**
	 * Return the code name of the sheet
	 *
	 * @return null|string
	*/
	public function getCodeName(){
		return $this->_codeName;
	}
	/**
	 * Sheet has a code name ?
	 * @return boolean
	*/
	public function hasCodeName(){
		return !(is_null($this->_codeName));
	}
}

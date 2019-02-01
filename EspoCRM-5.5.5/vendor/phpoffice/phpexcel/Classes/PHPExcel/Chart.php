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
 * @category	\PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package		\PhpOffice\PhpSpreadsheet\Chart\Chart
 * @copyright	Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		##VERSION##, ##DATE##
 */


/**
 * \PhpOffice\PhpSpreadsheet\Chart\Chart
 *
 * @category	\PhpOffice\PhpSpreadsheet\Spreadsheet
 * @package		\PhpOffice\PhpSpreadsheet\Chart\Chart
 * @copyright	Copyright (c) 2006 - 2014 \PhpOffice\PhpSpreadsheet\Spreadsheet (http://www.codeplex.com/\PhpOffice\PhpSpreadsheet\Spreadsheet)
 */
class \PhpOffice\PhpSpreadsheet\Chart\Chart
{
	/**
	 * Chart Name
	 *
	 * @var string
	 */
	private $_name = '';

	/**
	 * Worksheet
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 */
	private $_worksheet = null;

	/**
	 * Chart Title
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	private $_title = null;

	/**
	 * Chart Legend
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Chart\Legend
	 */
	private $_legend = null;

	/**
	 * X-Axis Label
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	private $_xAxisLabel = null;

	/**
	 * Y-Axis Label
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	private $_yAxisLabel = null;

	/**
	 * Chart Plot Area
	 *
	 * @var \PhpOffice\PhpSpreadsheet\Chart\PlotArea
	 */
	private $_plotArea = null;

	/**
	 * Plot Visible Only
	 *
	 * @var boolean
	 */
	private $_plotVisibleOnly = true;

	/**
	 * Display Blanks as
	 *
	 * @var string
	 */
	private $_displayBlanksAs = '0';

  /**
   * Chart Asix Y as
   *
   * @var \PhpOffice\PhpSpreadsheet\Chart\Axis
   */
  private $_yAxis = null;

  /**
   * Chart Asix X as
   *
   * @var \PhpOffice\PhpSpreadsheet\Chart\Axis
   */
  private $_xAxis = null;

  /**
   * Chart Major Gridlines as
   *
   * @var \PhpOffice\PhpSpreadsheet\Chart\GridLines
   */
  private $_majorGridlines = null;

  /**
   * Chart Minor Gridlines as
   *
   * @var \PhpOffice\PhpSpreadsheet\Chart\GridLines
   */
  private $_minorGridlines = null;

	/**
	 * Top-Left Cell Position
	 *
	 * @var string
	 */
	private $_topLeftCellRef = 'A1';


	/**
	 * Top-Left X-Offset
	 *
	 * @var integer
	 */
	private $_topLeftXOffset = 0;


	/**
	 * Top-Left Y-Offset
	 *
	 * @var integer
	 */
	private $_topLeftYOffset = 0;


	/**
	 * Bottom-Right Cell Position
	 *
	 * @var string
	 */
	private $_bottomRightCellRef = 'A1';


	/**
	 * Bottom-Right X-Offset
	 *
	 * @var integer
	 */
	private $_bottomRightXOffset = 10;


	/**
	 * Bottom-Right Y-Offset
	 *
	 * @var integer
	 */
	private $_bottomRightYOffset = 10;


	/**
	 * Create a new \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function __construct($name, \PhpOffice\PhpSpreadsheet\Chart\Title $title = null, \PhpOffice\PhpSpreadsheet\Chart\Legend $legend = null, \PhpOffice\PhpSpreadsheet\Chart\PlotArea $plotArea = null, $plotVisibleOnly = true, $displayBlanksAs = '0', \PhpOffice\PhpSpreadsheet\Chart\Title $xAxisLabel = null, \PhpOffice\PhpSpreadsheet\Chart\Title $yAxisLabel = null, \PhpOffice\PhpSpreadsheet\Chart\Axis $xAxis = null, \PhpOffice\PhpSpreadsheet\Chart\Axis $yAxis = null, \PhpOffice\PhpSpreadsheet\Chart\GridLines $majorGridlines = null, \PhpOffice\PhpSpreadsheet\Chart\GridLines $minorGridlines = null)
	{
		$this->_name = $name;
		$this->_title = $title;
		$this->_legend = $legend;
		$this->_xAxisLabel = $xAxisLabel;
		$this->_yAxisLabel = $yAxisLabel;
		$this->_plotArea = $plotArea;
		$this->_plotVisibleOnly = $plotVisibleOnly;
		$this->_displayBlanksAs = $displayBlanksAs;
		$this->_xAxis = $xAxis;
		$this->_yAxis = $yAxis;
    $this->_majorGridlines = $majorGridlines;
    $this->_minorGridlines = $minorGridlines;
	}

	/**
	 * Get Name
	 *
	 * @return string
	 */
	public function getName() {
		return $this->_name;
	}

	/**
	 * Get Worksheet
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
	 */
	public function getWorksheet() {
		return $this->_worksheet;
	}

	/**
	 * Set Worksheet
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet	$pValue
	 * @throws	\PhpOffice\PhpSpreadsheet\Chart\Exception
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setWorksheet(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $pValue = null) {
		$this->_worksheet = $pValue;

		return $this;
	}

	/**
	 * Get Title
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	public function getTitle() {
		return $this->_title;
	}

	/**
	 * Set Title
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Chart\Title $title
	 * @return	\PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setTitle(\PhpOffice\PhpSpreadsheet\Chart\Title $title) {
		$this->_title = $title;

		return $this;
	}

	/**
	 * Get Legend
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Legend
	 */
	public function getLegend() {
		return $this->_legend;
	}

	/**
	 * Set Legend
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Chart\Legend $legend
	 * @return	\PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setLegend(\PhpOffice\PhpSpreadsheet\Chart\Legend $legend) {
		$this->_legend = $legend;

		return $this;
	}

	/**
	 * Get X-Axis Label
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	public function getXAxisLabel() {
		return $this->_xAxisLabel;
	}

	/**
	 * Set X-Axis Label
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Chart\Title $label
	 * @return	\PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setXAxisLabel(\PhpOffice\PhpSpreadsheet\Chart\Title $label) {
		$this->_xAxisLabel = $label;

		return $this;
	}

	/**
	 * Get Y-Axis Label
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Title
	 */
	public function getYAxisLabel() {
		return $this->_yAxisLabel;
	}

	/**
	 * Set Y-Axis Label
	 *
	 * @param	\PhpOffice\PhpSpreadsheet\Chart\Title $label
	 * @return	\PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setYAxisLabel(\PhpOffice\PhpSpreadsheet\Chart\Title $label) {
		$this->_yAxisLabel = $label;

		return $this;
	}

	/**
	 * Get Plot Area
	 *
	 * @return \PhpOffice\PhpSpreadsheet\Chart\PlotArea
	 */
	public function getPlotArea() {
		return $this->_plotArea;
	}

	/**
	 * Get Plot Visible Only
	 *
	 * @return boolean
	 */
	public function getPlotVisibleOnly() {
		return $this->_plotVisibleOnly;
	}

	/**
	 * Set Plot Visible Only
	 *
	 * @param boolean $plotVisibleOnly
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setPlotVisibleOnly($plotVisibleOnly = true) {
		$this->_plotVisibleOnly = $plotVisibleOnly;

		return $this;
	}

	/**
	 * Get Display Blanks as
	 *
	 * @return string
	 */
	public function getDisplayBlanksAs() {
		return $this->_displayBlanksAs;
	}

	/**
	 * Set Display Blanks as
	 *
	 * @param string $displayBlanksAs
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setDisplayBlanksAs($displayBlanksAs = '0') {
		$this->_displayBlanksAs = $displayBlanksAs;
	}


  /**
   * Get yAxis
   *
   * @return \PhpOffice\PhpSpreadsheet\Chart\Axis
   */
  public function getChartAxisY() {
    if($this->_yAxis !== NULL){
      return $this->_yAxis;
    }

    return new \PhpOffice\PhpSpreadsheet\Chart\Axis();
  }

  /**
   * Get xAxis
   *
   * @return \PhpOffice\PhpSpreadsheet\Chart\Axis
   */
  public function getChartAxisX() {
    if($this->_xAxis !== NULL){
      return $this->_xAxis;
    }

    return new \PhpOffice\PhpSpreadsheet\Chart\Axis();
  }

  /**
   * Get Major Gridlines
   *
   * @return \PhpOffice\PhpSpreadsheet\Chart\GridLines
   */
  public function getMajorGridlines() {
    if($this->_majorGridlines !== NULL){
      return $this->_majorGridlines;
    }

    return new \PhpOffice\PhpSpreadsheet\Chart\GridLines();
  }

  /**
   * Get Minor Gridlines
   *
   * @return \PhpOffice\PhpSpreadsheet\Chart\GridLines
   */
  public function getMinorGridlines() {
    if($this->_minorGridlines !== NULL){
      return $this->_minorGridlines;
    }

    return new \PhpOffice\PhpSpreadsheet\Chart\GridLines();
  }


	/**
	 * Set the Top Left position for the chart
	 *
	 * @param	string	$cell
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setTopLeftPosition($cell, $xOffset=null, $yOffset=null) {
		$this->_topLeftCellRef = $cell;
		if (!is_null($xOffset))
			$this->setTopLeftXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setTopLeftYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the top left position of the chart
	 *
	 * @return array	an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
	 */
	public function getTopLeftPosition() {
		return array( 'cell'	=> $this->_topLeftCellRef,
					  'xOffset'	=> $this->_topLeftXOffset,
					  'yOffset'	=> $this->_topLeftYOffset
					);
	}

	/**
	 * Get the cell address where the top left of the chart is fixed
	 *
	 * @return string
	 */
	public function getTopLeftCell() {
		return $this->_topLeftCellRef;
	}

	/**
	 * Set the Top Left cell position for the chart
	 *
	 * @param	string	$cell
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setTopLeftCell($cell) {
		$this->_topLeftCellRef = $cell;

		return $this;
	}

	/**
	 * Set the offset position within the Top Left cell for the chart
	 *
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setTopLeftOffset($xOffset=null,$yOffset=null) {
		if (!is_null($xOffset))
			$this->setTopLeftXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setTopLeftYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the offset position within the Top Left cell for the chart
	 *
	 * @return integer[]
	 */
	public function getTopLeftOffset() {
		return array( 'X' => $this->_topLeftXOffset,
					  'Y' => $this->_topLeftYOffset
					);
	}

	public function setTopLeftXOffset($xOffset) {
		$this->_topLeftXOffset = $xOffset;

		return $this;
	}

	public function getTopLeftXOffset() {
		return $this->_topLeftXOffset;
	}

	public function setTopLeftYOffset($yOffset) {
		$this->_topLeftYOffset = $yOffset;

		return $this;
	}

	public function getTopLeftYOffset() {
		return $this->_topLeftYOffset;
	}

	/**
	 * Set the Bottom Right position of the chart
	 *
	 * @param	string	$cell
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setBottomRightPosition($cell, $xOffset=null, $yOffset=null) {
		$this->_bottomRightCellRef = $cell;
		if (!is_null($xOffset))
			$this->setBottomRightXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setBottomRightYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the bottom right position of the chart
	 *
	 * @return array	an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
	 */
	public function getBottomRightPosition() {
		return array( 'cell'	=> $this->_bottomRightCellRef,
					  'xOffset'	=> $this->_bottomRightXOffset,
					  'yOffset'	=> $this->_bottomRightYOffset
					);
	}

	public function setBottomRightCell($cell) {
		$this->_bottomRightCellRef = $cell;

		return $this;
	}

	/**
	 * Get the cell address where the bottom right of the chart is fixed
	 *
	 * @return string
	 */
	public function getBottomRightCell() {
		return $this->_bottomRightCellRef;
	}

	/**
	 * Set the offset position within the Bottom Right cell for the chart
	 *
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return \PhpOffice\PhpSpreadsheet\Chart\Chart
	 */
	public function setBottomRightOffset($xOffset=null,$yOffset=null) {
		if (!is_null($xOffset))
			$this->setBottomRightXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setBottomRightYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the offset position within the Bottom Right cell for the chart
	 *
	 * @return integer[]
	 */
	public function getBottomRightOffset() {
		return array( 'X' => $this->_bottomRightXOffset,
					  'Y' => $this->_bottomRightYOffset
					);
	}

	public function setBottomRightXOffset($xOffset) {
		$this->_bottomRightXOffset = $xOffset;

		return $this;
	}

	public function getBottomRightXOffset() {
		return $this->_bottomRightXOffset;
	}

	public function setBottomRightYOffset($yOffset) {
		$this->_bottomRightYOffset = $yOffset;

		return $this;
	}

	public function getBottomRightYOffset() {
		return $this->_bottomRightYOffset;
	}


	public function refresh() {
		if ($this->_worksheet !== NULL) {
			$this->_plotArea->refresh($this->_worksheet);
		}
	}

	public function render($outputDestination = null) {
		$libraryName = \PhpOffice\PhpSpreadsheet\Settings::getChartRendererName();
		if (is_null($libraryName)) {
			return false;
		}
		//	Ensure that data series values are up-to-date before we render
		$this->refresh();

		$libraryPath = \PhpOffice\PhpSpreadsheet\Settings::getChartRendererPath();
		$includePath = str_replace('\\','/',get_include_path());
		$rendererPath = str_replace('\\','/',$libraryPath);
		if (strpos($rendererPath,$includePath) === false) {
			set_include_path(get_include_path() . PATH_SEPARATOR . $libraryPath);
		}

		$rendererName = '\PhpOffice\PhpSpreadsheet\Chart\Chart_Renderer_'.$libraryName;
		$renderer = new $rendererName($this);

		if ($outputDestination == 'php://output') {
			$outputDestination = null;
		}
		return $renderer->render($outputDestination);
	}

}

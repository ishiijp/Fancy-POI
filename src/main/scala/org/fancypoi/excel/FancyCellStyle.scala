package org.fancypoi.excel

import org.apache.poi.ss.usermodel.{IndexedColors, Font, Color, CellStyle}

/**
 * CellStyleのバリューオブジェクト
 */
object FancyCellStyle {
	val DEFAULT_CELL_STYLE_INDEX = -1 toShort
}

class FancyCellStyle extends CellStyle {
	private var _align: Short = 0
	private var _backgroundColor: Short = IndexedColors.WHITE.getIndex
	private var _borderBottom: Short = CellStyle.BORDER_NONE
	private var _borderLeft: Short = CellStyle.BORDER_NONE
	private var _borderRight: Short = CellStyle.BORDER_NONE
	private var _borderTop: Short = CellStyle.BORDER_NONE
	private var _bottomBorderColor: Short = IndexedColors.BLACK.getIndex
	private var _dataFormat: Short = 0
	private var _fillPattern: Short = CellStyle.NO_FILL
	private var _font: Font = _
	private var _foregroundColor: Short = IndexedColors.WHITE.getIndex
	private var _hidden: Boolean = false
	private var _indent: Short = 0
	private var _leftBorderColor: Short = IndexedColors.BLACK.getIndex
	private var _locked: Boolean = true
	private var _rightBorderColor: Short = IndexedColors.BLACK.getIndex
	private var _rotation: Short = 0
	private var _topBorderColor: Short = IndexedColors.BLACK.getIndex
	private var _verticalAlign: Short = 2
	private var _wrapped: Boolean = false


	def getIndex: Short = FancyCellStyle.DEFAULT_CELL_STYLE_INDEX

	def getAlignment: Short = _align

	def setAlignment(align: Short): Unit = _align = align

	def getBorderBottom: Short = _borderBottom

	def setBorderBottom(border: Short): Unit = _borderBottom = border

	def getBorderLeft: Short = _borderLeft

	def setBorderLeft(border: Short): Unit = _borderLeft = border

	def getBorderRight: Short = _borderRight

	def setBorderRight(border: Short): Unit = _borderRight = border

	def getBorderTop: Short = _borderTop

	def setBorderTop(border: Short): Unit = _borderTop = border

	def getBottomBorderColor: Short = _bottomBorderColor

	def setBottomBorderColor(color: Short): Unit = _bottomBorderColor = color

	def getDataFormat: Short = _dataFormat

	def getDataFormatString: String = throw new RuntimeException("Can't stringize data format.")

	def setDataFormat(fmt: Short): Unit = _dataFormat = fmt

	def getFillBackgroundColor: Short = _backgroundColor

	def getFillBackgroundColorColor: Color = throw new RuntimeException("Can't convert color index to color object.")

	def setFillBackgroundColor(bg: Short): Unit = _backgroundColor = bg

	def getFillForegroundColor: Short = _foregroundColor

	def getFillForegroundColorColor: Color = throw new RuntimeException("Can't convert color index to color object.")

	def setFillForegroundColor(fg: Short): Unit = _foregroundColor = fg

	def getFillPattern: Short = _fillPattern

	def setFillPattern(fp: Short): Unit = _fillPattern = fp

	def getFontIndex: Short = _font.getIndex

	def setFont(font: Font): Unit = _font = font

	// CellStyleVOの固有メソッド
	def getFont = _font

	def getHidden: Boolean = _hidden

	def setHidden(hidden: Boolean): Unit = _hidden = hidden

	def getIndention: Short = _indent

	def setIndention(indent: Short): Unit = _indent = indent

	def getLeftBorderColor: Short = _leftBorderColor

	def setLeftBorderColor(color: Short): Unit = _leftBorderColor = color

	def getLocked: Boolean = _locked

	def setLocked(locked: Boolean): Unit = _locked = locked

	def getRightBorderColor: Short = _rightBorderColor

	def setRightBorderColor(color: Short): Unit = _rightBorderColor = color

	def getRotation: Short = _rotation

	def setRotation(rotation: Short): Unit = _rotation = rotation

	def getTopBorderColor: Short = _topBorderColor

	def setTopBorderColor(color: Short): Unit = _topBorderColor = color

	def getVerticalAlignment: Short = _verticalAlign

	def setVerticalAlignment(align: Short): Unit = _verticalAlign = align

	def getWrapText: Boolean = _wrapped

	def setWrapText(wrapped: Boolean): Unit = _wrapped = wrapped

	def cloneStyleFrom(source: CellStyle): Unit = throw new RuntimeException("Can't clone style.")
}


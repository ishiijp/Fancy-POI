package org.fancypoi.excel

import org.apache.poi.ss.usermodel.Font

object FancyFont {
	val DEFAULT_FONT_INDEX = -1 toShort
}

class FancyFont extends Font {
	def getIndex: Short = FancyFont.DEFAULT_FONT_INDEX

	private var _name: String = "Arial"
	private var _fontHeight: Short = 200
	private var _italic: Boolean = false
	private var _strikeout: Boolean = false
	private var _color: Short = 0x7fff toShort
	private var _offset: Short = 0
	private var _underline: Byte = Font.U_NONE
	private var _charset: Int = 0
	private var _boldweight: Short = 400

	def setFontName(name: String): Unit = _name = name

	def getFontName: String = _name

	def setFontHeight(height: Short): Unit = _fontHeight = height

	def setFontHeightInPoints(height: Short): Unit = _fontHeight = height * 20 toShort

	def getFontHeight: Short = _fontHeight

	def getFontHeightInPoints: Short = _fontHeight * 20 toShort

	def setItalic(italic: Boolean): Unit = _italic = italic

	def getItalic: Boolean = _italic

	def setStrikeout(strikeout: Boolean): Unit = _strikeout = strikeout

	def getStrikeout: Boolean = _strikeout

	def setColor(color: Short): Unit = _color = color

	def getColor: Short = _color

	def setTypeOffset(offset: Short): Unit = _offset = offset

	def getTypeOffset: Short = _offset

	def setUnderline(underline: Byte): Unit = _underline = underline

	def getUnderline: Byte = _underline

	def setCharSet(charset: Byte): Unit = _charset = charset toInt

	def setCharSet(charset: Int): Unit = _charset = charset

	def getCharSet: Int = _charset

	def setBoldweight(boldweight: Short): Unit = _boldweight = boldweight

	def getBoldweight: Short = _boldweight
}

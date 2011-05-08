package org.fancypoi.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io._
import FancyExcelUtils._
import org.apache.poi.ss.usermodel._

/**
 * User: ishiiyoshinori
 * Date: 11/05/04
 */

object FancyWorkbook {

	def createXls = new HSSFWorkbook

	def createXlsx = new XSSFWorkbook

	def createFromFile(file: File): Workbook = {
		val fis = new FileInputStream(file)
		val w = WorkbookFactory.create(fis)
		fis.close
		w
	}

	def createFromInputStream(is: InputStream) = {
		val w = WorkbookFactory.create(is)
		is.close
		w
	}
}


class FancyWorkbook(protected[fancypoi] val workbook: Workbook) {
	override def toString = "Workbook"

	protected[fancypoi] val defaultFont = new FancyFont

	protected[fancypoi] val defaultStyle = {
		val style = new FancyCellStyle
		style.setFont(defaultFont)
		style
	}

	protected[fancypoi] val tmpFont = new FancyFont

	protected[fancypoi] val tmpStyle = new FancyCellStyle

	def getFontAt(index: Short) = index == FancyFont.DEFAULT_FONT_INDEX match {
		case true => defaultFont
		case false => workbook.getFontAt(index)
	}

	def getStyleAt(index: Short) = index == FancyCellStyle.DEFAULT_CELL_STYLE_INDEX match {
		case true => defaultStyle
		case false => workbook.getCellStyleAt(index)
	}

	/**
	 * blockで設定したフォントを取得します。
	 * 既に同じ値を持つフォントがある場合は、それを返し新しいフォントは生成しません。
	 */
	def getFontWith(block: Font => Unit): Font = getFontBasedWith(defaultFont)(block)

	/**
	 * ベースとなるフォントを指定し、blockで設定したフォントを取得します。
	 * 既に同じ値を持つフォントがある場合は、それを返し新しいフォントは生成しません。
	 */
	def getFontBasedWith(base: Font)(block: Font => Unit): Font = {
		copyFont(base, tmpFont) // デフォルトフォントを一時フォントにコピー
		block(tmpFont) // 一時フォントを設定
		searchFont(this, tmpFont) match {
			case Some(font) => font // フォントがすでにある場合はそれを返す
			case None =>
				val newFont = workbook.createFont
				copyFont(tmpFont, newFont)
				newFont
		}
	}

	/**
	 * blockで設定したセルスタイルを取得します。
	 * 既に同じ値を持つセルスタイルがある場合は、それを返し新しいセルスタイルは生成しません。
	 */
	def getStyle(block: CellStyle => Unit): CellStyle = getStyleBasedWith(defaultStyle)(block)

	/**
	 * ベースとなるセルスタイルを指定し、blockで設定したスタイルを取得します。
	 * 既に同じ値を持つセルスタイルがある場合は、それを返し新しいセルスタイルは生成しません。
	 */
	def getStyleBasedWith(base: CellStyle)(block: CellStyle => Unit) = {

		// tmpStyle と tmpFont をベースとなるセルスタイルで初期化
		copyStyleWithoutFont(base, tmpStyle)
		copyFont(getFontAt(base.getFontIndex), tmpFont)
		tmpStyle.setFont(tmpFont)

		// スタイルを設定
		block(tmpStyle)

		// ワークブックから取得したいセルスタイルを検索し、ない場合は生成する。
		searchStyle(this, tmpStyle) match {
			case Some(style) => style
			case None =>
				val newStyle = workbook.createCellStyle
				copyStyleWithoutFont(tmpStyle, newStyle)
				val font = tmpStyle.getFontIndex == tmpFont.getIndex match {
					case true =>
						val f = workbook.createFont
						copyFont(tmpFont, f)
						f
					case false => workbook.getFontAt(tmpStyle.getFontIndex)
				}
				newStyle.setFont(font)
				newStyle
		}
	}

	/**
	 * シートを名前で検索します
	 */
	def sheet(name: String) = !!(workbook.getSheet(name)) match {
		case None => workbook.createSheet(name)
		case Some(sheet) => sheet
	}

	def sheet_?(name: String): Option[Sheet] = !!(workbook.getSheet(name))

	/**
	 * シートをインデックスで検索し、blockを適用します。
	 */
	def sheetAt(index: Int) = workbook.getSheetAt(index)


	def sheetAt_?(index: Int): Option[Sheet] = workbook.getNumberOfSheets - 1 < index match {
		case true => None
		case false => Some(workbook.getSheetAt(index))
	}

	/**
	 * シートをリストで返します。
	 */
	def sheets = (0 to workbook.getNumberOfSheets - 1).map(sheetAt).toList


	/**
	 *  ワークブックをファイルに書き出します。
	 */
	def write(file: File) {
		val fos = new FileOutputStream(file)
		val bos = new BufferedOutputStream(fos)
		workbook.write(bos)
		bos.close
	}
}
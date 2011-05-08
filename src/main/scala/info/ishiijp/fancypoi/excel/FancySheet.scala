package org.fancypoi.excel

import org.apache.poi.ss.usermodel.{Sheet, Row, Cell}
import FancyExcelUtils._
import org.fancypoi.Implicits._

/**
 * User: ishiiyoshinori
 * Date: 11/05/04
 */

class FancySheet(protected[fancypoi] val _sheet: Sheet) {
	lazy val workbook = _sheet.getWorkbook

	override def toString = "#" + _sheet.getSheetName

	def cell(address: String): Cell = {
		val (colIndex, rowIndex) = addrToIndexes(address)
		cellAt(colIndex, rowIndex)
	}

	def cellAt(colIndex: Int, rowIndex: Int): Cell = {
		rowAt(rowIndex).cellAt(colIndex)
	}

	def cell_?(address: String): Option[Cell] = {
		val (colIndex, rowIndex) = addrToIndexes(address)
		cellAt_?(colIndex, rowIndex)
	}

	def cellAt_?(colIndex: Int, rowIndex: Int): Option[Cell] = {
		rowAt_?(rowIndex).flatMap(r => r.cellAt_?(colIndex))
	}

	def row(address: String): Row = rowAt(address.toInt - 1)

	def rowAt(index: Int): Row = rowAt_?(index) match {
		case Some(row) => row
		case None => _sheet.createRow(index)
	}

	def row_?(address: String): Option[Row] = rowAt_?(address.toInt - 1)

	def rowAt_?(index: Int): Option[Row] = !!(_sheet.getRow(index))

	def rows: List[Row] = (0 to lastRowIndex).map(rowAt).toList

	def rows(rowRange: (String, String)): List[Row] = {
		val startIndex = rowRange._1.toInt - 1
		val endIndex = rowRange._2.toInt - 1
		(startIndex to endIndex).map(rowAt).toList
	}

	def rowsAt(rowIndexRange: (Int, Int)) = {
		(rowIndexRange._1 to rowIndexRange._2).map(rowAt).toList
	}

	def insertRows(rowAddr: String, num: Int) = insertRowsAt(rowAddr.toInt - 1, num)

	def insertRowsAt(insertRowIndex: Int, num: Int) {
		if (0 < num) {
			val endRowIndex = _sheet.getLastRowNum < insertRowIndex match {
				case true => insertRowIndex
				case false => _sheet.getLastRowNum
			}
			_sheet.shiftRows(insertRowIndex, endRowIndex, num, true, true)
		}
	}

	def firstRowIndex = _sheet.getTopRow

	def firstRowAddr = _sheet.getTopRow + 1 toString

	def lastRowIndex = _sheet.getLastRowNum

	def lastRowAddr = _sheet.getLastRowNum + 1 toString
}
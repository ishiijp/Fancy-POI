package org.fancypoi.excel

import org.apache.poi.ss.usermodel.{Row, Cell}
import FancyExcelUtils._
import org.fancypoi.Implicits._

class FancyRow(protected[fancypoi] val _row: Row) {

	override def toString = "#" + _row.getSheet.getSheetName + "!*" + addr

	def addr = (_row.getRowNum + 1).toString

	def cell(address: String): Cell = cellAt(colAddrToIndex(address))

	def cellAt(index: Int) = _row.getCell(index, Row.CREATE_NULL_AS_BLANK)

	def cell_?(address:String) = cellAt_?(colAddrToIndex(address))

	def cellAt_?(index:Int) = !!(_row.getCell(index, Row.RETURN_NULL_AND_BLANK))

	def cells: List[Cell] = (0 to lastColIndex).map(cellAt).toList

	def firstColAddr = colIndexToAddr(firstColIndex)

	def firstColIndex = _row.getFirstCellNum.toInt

	def lastColAddr = colIndexToAddr(lastColIndex)

	def lastColIndex = _row.getLastCellNum.toInt

	def cellsFrom(startColAddr: String)(block: CellSeq => Unit) {
		cellsFromAt(colAddrToIndex(startColAddr))(block)
	}

	def cellsFromAt(startColIndex: Int)(block: CellSeq => Unit) {
		block(new CellSeq(_row, startColIndex))
	}

	private class CellSeq(row: Row, colIndex: Int) {
		var current = colIndex

		def apply(block: Cell => Unit) {
			block(row.cellAt(current))
			current += 1
		}
	}

}
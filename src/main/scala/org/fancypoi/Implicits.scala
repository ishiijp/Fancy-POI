package org.fancypoi

import excel.FancyExcelUtils.AddrRangeStart
import excel.{FancyWorkbook, FancySheet, FancyRow, FancyCell}
import org.apache.poi.ss.usermodel._

/**
 * User: ishiiyoshinori
 * Date: 11/05/04
 */

object Implicits {

	implicit def workbook2fancy(w: Workbook) = new FancyWorkbook(w)

	implicit def sheet2fancy(s: Sheet) = new FancySheet(s)

	implicit def row2fancy(r: Row) = new FancyRow(r)

	implicit def cell2fancy(c: Cell) = new FancyCell(c)

	implicit def workbook2plain(w: FancyWorkbook) = w.workbook

	implicit def sheet2plain(s: FancySheet) = s._sheet

	implicit def row2plain(r: FancyRow) = r._row

	implicit def cell2plain(c: FancyCell) = c._cell

	implicit def indexedColors2Int(indexedColor: IndexedColors) = indexedColor.getIndex

	implicit def str2Addr(addr: String) = new AddrRangeStart(addr)

}
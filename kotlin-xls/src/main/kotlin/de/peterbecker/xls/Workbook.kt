package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.ss.util.CellReference

fun Workbook.writeToRange(name: String, rows: Iterator<List<Any?>>) {
    val range = this.getName(name) ?: throw NamedRangeNotFound(name)
    val ref = AreaReference(range.refersToFormula, SpreadsheetVersion.EXCEL2007)
    val sheet = this.getSheet(range.sheetName)
    val width = ref.lastCell.col - ref.firstCell.col + 1
    var r = 0
    for (row in rows) {
        if (row.size > width) {
            throw RowTooLongException(name, r + 1, width, row.size)
        }
        for ((c, value) in row.withIndex()) {
            sheet.setValueAt(ref.firstCell.row + r, ref.firstCell.col + c, value)
        }
        r++
    }
    range.refersToFormula = AreaReference(
        ref.firstCell,
        CellReference(ref.firstCell.row + r, ref.lastCell.col),
        SpreadsheetVersion.EXCEL2007
    ).formatAsString()
}

fun Workbook.writeToRange(name: String, rows: Iterable<List<Any?>>) {
    writeToRange(name, rows.iterator())
}
package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.usermodel.Name
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.ss.util.CellReference

fun Workbook.writeToRange(name: String, rows: Iterator<List<Any?>>): AreaReference {
    val range = this.getName(name) ?: throw NamedRangeNotFound(name)
    return writeToRange(range, rows)
}

fun Workbook.writeToRange(range: Name, rows: Iterator<List<Any?>>): AreaReference {
    val ref = AreaReference(range.refersToFormula, SpreadsheetVersion.EXCEL2007)
    val sheet = this.getSheet(range.sheetName)
    val newArea = sheet.writeToArea(ref, rows)
    range.refersToFormula = newArea.formatAsString()
    return newArea
}

fun Workbook.writeToRange(name: String, rows: Iterable<List<Any?>>) =
        writeToRange(name, rows.iterator())

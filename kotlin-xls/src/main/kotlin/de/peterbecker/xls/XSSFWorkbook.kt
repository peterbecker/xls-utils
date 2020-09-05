package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.IOException

fun XSSFWorkbook.writeToTable(name: String, rows: Iterator<List<Any?>>) {
    val table = sheetIterator().asSequence().mapNotNull {sh ->
        when (sh) {
            is XSSFSheet -> sh.tables.firstOrNull { it.name == name }
            else -> throw IOException("Invalid input file: XSSF workbook contains non-XSSF sheet")
        }
    }.firstOrNull()?: throw NamedTableNotFound(name)
    val areaNoHeader = AreaReference(
        CellReference(table.area.firstCell.row + 1, table.area.firstCell.col),
        CellReference(table.area.lastCell.row, table.area.lastCell.col),
        SpreadsheetVersion.EXCEL2007
    )
    val r = table.xssfSheet.writeToArea(areaNoHeader, rows)
    if(r > 0) {
        table.area = AreaReference(
            table.area.firstCell,
            CellReference(table.area.firstCell.row + r, table.area.lastCell.col),
            SpreadsheetVersion.EXCEL2007
        )
    }
}

fun XSSFWorkbook.writeToTable(name: String, rows: Iterable<List<Any?>>) {
    writeToTable(name, rows.iterator())
}

fun XSSFWorkbook.writeData(targetName: String, rows: Iterator<List<Any?>>) {
    when (val nameObject = this.getName(targetName)) {
        null -> writeToTable(targetName, rows)
        else -> writeToRange(nameObject, rows)
    }
}

fun XSSFWorkbook.writeData(targetName: String, rows: Iterable<List<Any?>>) {
    writeData(targetName, rows.iterator())
}

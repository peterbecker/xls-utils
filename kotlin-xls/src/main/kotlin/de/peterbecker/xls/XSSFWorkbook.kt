package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.IOException

fun XSSFWorkbook.writeToTable(name: String, rows: Iterator<List<Any?>>) {
    val table = sheetIterator().asSequence().mapNotNull { sh ->
        when (sh) {
            is XSSFSheet -> sh.tables.firstOrNull { it.name == name }
            else -> throw IOException("Invalid input file: XSSF workbook contains non-XSSF sheet")
        }
    }.firstOrNull() ?: throw NamedTableNotFound(name)
    val areaNoHeader = table.area.derive(top = 1)
    table.area = table.xssfSheet.writeToArea(areaNoHeader, rows).derive(top = -1)
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

fun XSSFWorkbook.findChartByTitle(title: String) =
        this.sheetIterator().asSequence()
                .mapNotNull {
                    when (it) {
                        is XSSFSheet -> it.findChartByTitle(title)
                        else -> null
                    }
                }
                .firstOrNull()
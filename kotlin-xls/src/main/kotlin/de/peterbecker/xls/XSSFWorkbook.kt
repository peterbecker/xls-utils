package de.peterbecker.xls

import org.apache.poi.ss.util.AreaReference
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.IOException

/**
 * Updates the table to contain the data provided. Returns the new data range, excluding the header.
 */
fun XSSFWorkbook.writeToTable(name: String, rows: Iterator<List<Any?>>): AreaReference {
    val table = sheetIterator().asSequence().mapNotNull { sh ->
        when (sh) {
            is XSSFSheet -> sh.tables.firstOrNull { it.name == name }
            else -> throw IOException("Invalid input file: XSSF workbook contains non-XSSF sheet")
        }
    }.firstOrNull() ?: throw NamedTableNotFound(name)
    val areaNoHeader = table.area.derive(top = 1)
    val newDataArea = table.xssfSheet.writeToArea(areaNoHeader, rows)
    table.area = newDataArea.derive(top = -1)
    return newDataArea
}

fun XSSFWorkbook.writeToTable(name: String, rows: Iterable<List<Any?>>): AreaReference {
    return writeToTable(name, rows.iterator())
}

fun XSSFWorkbook.writeData(targetName: String, rows: Iterator<List<Any?>>) =
    when (val nameObject = this.getName(targetName)) {
        null -> writeToTable(targetName, rows)
        else -> writeToRange(nameObject, rows)
    }

fun XSSFWorkbook.writeData(targetName: String, rows: Iterable<List<Any?>>) =
    writeData(targetName, rows.iterator())

fun XSSFWorkbook.findChartByTitle(title: String) =
        this.sheetIterator().asSequence()
                .mapNotNull {
                    when (it) {
                        is XSSFSheet -> it.findChartByTitle(title)
                        else -> null
                    }
                }
                .firstOrNull()

/**
 * Expands the references charts have to data ranges if the existing reference is within the given area, in a way that
 * the bottom matches the new area.
 *
 * More accurately: if the chart refers to an area that is in the bounds of the provided one, and the top of the areas
 * match, then the chart's area reference will be updated to be to the bottom of the provided one.
 */
fun XSSFWorkbook.expandChartReferences(area: AreaReference) {
    this.sheetIterator().forEach {
        when (it) {
            is XSSFSheet -> it.expandChartReferences(area)
            // we are not expected another case here
        }
    }
}

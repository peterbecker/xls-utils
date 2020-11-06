package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.ss.util.CellReference
import java.awt.Dimension

fun Sheet.getValueAt(row: Int, col: Int) = this.getRow(row)?.getCell(col)?.getValue()

fun Sheet.setValueAt(row: Int, col: Int, value: Any?) {
    when (value) {
        null -> {
            val r = this.getRow(row)
            if(r != null) {
                val cell = r.getCell(col)
                if(cell != null) {
                    r.removeCell(cell)
                }
            }
        }
        else -> this.getOrCreateRow(row).getOrCreateCell(col).setValue(value)
    }
}

fun Sheet.getOrCreateRow(row: Int) = this.getRow(row) ?: this.createRow(row)!!

/**
 * Writes data into the provided area and beyond, returning the area used when writing.
 */
fun Sheet.writeToArea(ref: AreaReference, rows: Iterator<List<Any?>>): AreaReference {
    val width = ref.lastCell.col - ref.firstCell.col + 1
    var r = 0
    for (row in rows) {
        if (row.size > width) {
            throw RowTooLongException(r, width, row.size)
        }
        for ((c, value) in row.withIndex()) {
            setValueAt(ref.firstCell.row + r, ref.firstCell.col + c, value)
        }
        r++
    }
    return AreaReference(
            ref.firstCell, ref.lastCell.replace(row = ref.firstCell.row + r - 1), SpreadsheetVersion.EXCEL2007
    )
}
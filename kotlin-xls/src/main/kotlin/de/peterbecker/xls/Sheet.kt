package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.AreaReference
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

fun Sheet.writeToArea(ref: AreaReference, rows: Iterator<List<Any?>>): Int {
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
    return r
}
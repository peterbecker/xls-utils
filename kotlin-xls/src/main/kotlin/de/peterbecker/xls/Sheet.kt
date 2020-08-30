package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet

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

package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType

fun Cell.getValue(): Any? =
    when (this.cellType) {
        CellType.BLANK -> null
        CellType.BOOLEAN -> this.booleanCellValue
        CellType.STRING -> this.stringCellValue
        CellType.NUMERIC -> this.numericCellValue
        else -> null // we still don't support all types
    }

fun Cell.setValue(value: Any) {
    when (value) {
        is Boolean -> {
            this.setCellValue(value)
        }
        is Number -> {
            this.setCellValue(value.toDouble())
        }
        else -> {
            this.setCellValue(value.toString())
        }
    }
}

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

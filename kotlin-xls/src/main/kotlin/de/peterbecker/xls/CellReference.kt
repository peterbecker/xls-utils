package de.peterbecker.xls

import org.apache.poi.ss.util.CellReference

/**
 * Returns a cell reference shifted by a number of rows/columns, maintaining the sheet and the absolute/relative states.
 */
fun CellReference.shift(row: Int = 0, col: Int = 0) =
        CellReference(this.sheetName, this.row + row, this.col + col, this.isRowAbsolute, this.isColAbsolute)

fun CellReference.replace(row: Int? = null, col: Int? = null) =
        CellReference(this.sheetName, row ?: this.row, col ?: this.col.toInt(), this.isRowAbsolute, this.isColAbsolute)
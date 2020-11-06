package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference


fun AreaReference.contains(other: AreaReference, ignoreSheet: Boolean = false) =
        (ignoreSheet || this.firstCell.sheetName == other.firstCell.sheetName) &&
                this.firstCell.col <= other.firstCell.col &&
                this.firstCell.row <= other.firstCell.row &&
                this.lastCell.col >= other.lastCell.col &&
                this.lastCell.row >= other.lastCell.row

val AreaReference.height: Int
    get() = this.lastCell.row - this.firstCell.row + 1

val AreaReference.width: Int
    get() = this.lastCell.col - this.firstCell.col + 1

/**
 * Creates an area reference that is slightly larger or smaller than the original.
 *
 * Note that there is currently no validation, which means that values could go out of range, but also that the cell
 * references can flip: if e.g. the left column gets changed to be past the right, then the roles will automatically
 * switch.
 */
fun AreaReference.derive(top: Int = 0, left: Int = 0, bottom: Int = 0, right: Int = 0) =
        AreaReference(
                this.firstCell.shift(top, left),
                this.lastCell.shift(bottom, right),
                SpreadsheetVersion.EXCEL2007
        )
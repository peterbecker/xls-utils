package de.peterbecker.xls

import org.apache.poi.ss.util.AreaReference


fun AreaReference.contains(other: AreaReference) =
        this.firstCell.sheetName == other.firstCell.sheetName &&
                this.firstCell.col <= other.firstCell.col &&
                this.firstCell.row <= other.firstCell.row &&
                this.lastCell.col >= other.lastCell.col &&
                this.lastCell.row >= other.lastCell.row

val AreaReference.height: Int
    get() = this.lastCell.row - this.firstCell.row + 1

val AreaReference.width: Int
    get() = this.lastCell.col - this.firstCell.col + 1

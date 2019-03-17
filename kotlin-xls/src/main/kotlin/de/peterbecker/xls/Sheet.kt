package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet

fun Sheet.getValueAt(row: Int, col: Int) = this.getRow(row)?.getCell(col)?.getValue()

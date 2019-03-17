package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet

fun Sheet.getValueAt(row: Int, col: Int) = this.getRow(row)?.getCell(col)?.getValue()

fun Sheet.getBooleanValueAt(row: Int, col: Int) = this.getRow(row)?.getCell(col)?.getValue() as Boolean?
fun Sheet.getDoubleValueAt(row: Int, col: Int) = this.getRow(row)?.getCell(col)?.getValue() as Double?
fun Sheet.getStringValueAt(row: Int, col: Int) = this.getRow(row)?.getCell(col)?.getValue() as String?

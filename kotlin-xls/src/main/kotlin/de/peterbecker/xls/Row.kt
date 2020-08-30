package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Row

fun Row.getOrCreateCell(col: Int) = this.getCell(col) ?: this.createCell(col)!!

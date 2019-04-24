package de.peterbecker.xls

import java.lang.IllegalArgumentException

class NoSuchRowException(
    val sheetName: String,
    val rowNumber: Int
) : IllegalArgumentException("Out of range: no row $rowNumber on sheet '$sheetName'")

class NoSuchCellException(
    val sheetName: String,
    val rowNumber: Int,
    val colNumber: Int
) : IllegalArgumentException("Out of range: no column $colNumber in row $rowNumber on sheet '$sheetName'")

class RowTooLongException(
    val targetRange: String,
    val rowNumber: Int,
    val rangeWidth: Int,
    val rowLength: Int
) : IllegalArgumentException("Row number $rowNumber for target range $targetRange is too long ($rowLength/$rangeWidth)")

class NamedRangeNotFound(
    val name: String
) : IllegalArgumentException("Could not find range named '$name'")

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
    val rowNumber: Int,
    val rangeWidth: Int,
    val rowLength: Int
) : IllegalArgumentException("Row number $rowNumber is too long ($rowLength/$rangeWidth) for target")

open class NameNotFound(
    open val name: String,
    val type: String
) : IllegalArgumentException("Could not find $type named '$name'")

class NamedRangeNotFound(override val name: String): NameNotFound(name, "range")
class NamedTableNotFound(override val name: String): NameNotFound(name, "table")


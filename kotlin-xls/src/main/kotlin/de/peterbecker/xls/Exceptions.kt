package de.peterbecker.xls

import java.lang.IllegalArgumentException

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


package de.peterbecker.xls

data class DiffMode(
    val allowExtraSheets: Boolean = false,
    val allowSheetsMissing: Boolean = false,
    val allowSheetNameDifference: Boolean = false
) {
    companion object {
        val Strict = DiffMode()
    }
}
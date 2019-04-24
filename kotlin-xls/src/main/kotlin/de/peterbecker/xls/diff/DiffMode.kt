package de.peterbecker.xls.diff

data class DiffMode(
    val allowExtraSheets: Boolean = false,
    val allowSheetsMissing: Boolean = false,
    val allowSheetNameDifference: Boolean = false
) {
    companion object {
        val Strict = DiffMode()
    }
}
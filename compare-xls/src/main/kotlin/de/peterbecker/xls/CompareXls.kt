package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellReference

fun compareWorkbooks(toCheck: Workbook, compareTo: Workbook, diffMode: DiffMode = DiffMode.Strict): ComparisonResult {
    val differences = ArrayList<Difference>()
    toCheck.sheetIterator().forEach {
        val compareToSheet = compareTo.getSheet(it.sheetName)
        when(compareToSheet) {
            null ->
                if(!diffMode.allowExtraSheets) {
                    differences.add(StructuralDifference(it.sheetName, "Extra sheet present: '${it.sheetName}'"))
                }
            else -> compareSheets(it, compareToSheet)
        }
    }
    if(!diffMode.allowSheetsMissing) {
        compareTo.sheetIterator().forEach {
            if (toCheck.getSheet(it.sheetName) == null) {
                differences.add(StructuralDifference(it.sheetName, "Sheet missing: '${it.sheetName}'"))
            }
        }
    }
    return if (differences.isEmpty()) Same else Different(differences)
}

fun compareSheets(toCheck: Sheet, compareTo: Sheet, diffMode: DiffMode = DiffMode.Strict): ComparisonResult {
    val differences = ArrayList<Difference>()
    if(!diffMode.allowSheetNameDifference && toCheck.sheetName != compareTo.sheetName) {
        differences.add(
            StructuralDifference(
                toCheck.sheetName,
                "Sheet names differ: '${toCheck.sheetName}' instead of '${compareTo.sheetName}'"
            )
        )
    }
    return if (differences.isEmpty()) Same else Different(differences)
}

sealed class ComparisonResult
object Same : ComparisonResult()
data class Different(val differences: List<Difference>) : ComparisonResult()

sealed class Difference() {
    abstract val message: String
}
data class StructuralDifference(val sheetName: String, override val message: String) : Difference()
data class CellContentDifference(val cell: CellReference, val left: Any, val right: Any, override val message: String) :
    Difference()

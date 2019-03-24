package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellReference

fun compareWorkbooks(toCheck: Workbook, compareTo: Workbook, diffMode: DiffMode = DiffMode.Strict): ComparisonResult {
    val differences = ArrayList<Difference>()
    toCheck.sheetIterator().forEach {
        val compareToSheet = compareTo.getSheet(it.sheetName)
        when (compareToSheet) {
            null ->
                if (!diffMode.allowExtraSheets) {
                    differences.add(StructuralDifference(it.sheetName, "Extra sheet present: '${it.sheetName}'"))
                }
            else -> {
                val sheetCompare = compareSheets(it, compareToSheet)
                when (sheetCompare) {
                    is Different -> differences.addAll(sheetCompare.differences)
                }
            }
        }
    }
    if (!diffMode.allowSheetsMissing) {
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
    if (!diffMode.allowSheetNameDifference && toCheck.sheetName != compareTo.sheetName) {
        differences.add(
            StructuralDifference(
                toCheck.sheetName,
                "Sheet names differ: '${toCheck.sheetName}' instead of '${compareTo.sheetName}'"
            )
        )
    }
    for (r in toCheck.firstRowNum..toCheck.lastRowNum) {
        val rowToCheck = toCheck.getRow(r)
        val rowToCompare = compareTo.getRow(r)
        when (rowToCheck) {
            null -> if (rowToCompare != null) {
                differences.add(
                    StructuralDifference(
                        toCheck.sheetName,
                        "Extra row $r in sheet '${toCheck.sheetName}'"
                    )
                )
            }
            else -> when (rowToCompare) {
                null -> {
                    differences.add(
                        StructuralDifference(
                            toCheck.sheetName,
                            "Missing row $r in sheet '${toCheck.sheetName}'"
                        )
                    )
                }
                else ->
                    differences.addAll(compareRows(rowToCheck, rowToCompare))
            }
        }
    }
    return if (differences.isEmpty()) Same else Different(differences)
}

fun compareRows(rowToCheck: Row, rowToCompare: Row): Collection<Difference> {
    val differences = ArrayList<Difference>()
    for (c in rowToCheck.firstCellNum..rowToCheck.lastCellNum) {
        val cellToCheck = rowToCheck.getCell(c)
        val cellToCompare = rowToCompare.getCell(c)
        when (cellToCheck) {
            null -> if (cellToCompare != null) {
                differences.add(
                    CellContentDifference(
                        CellReference(rowToCheck.sheet.sheetName, rowToCheck.rowNum, c, false, false),
                        null,
                        cellToCompare.getValue(),
                        "Cell ${crs(rowToCheck, c)} has no value when it should be ${v(cellToCompare)}"
                    )
                )
            }
            else -> when (cellToCompare) {
                null -> differences.add(
                    CellContentDifference(
                        cr(rowToCheck, c),
                        cellToCheck.getValue(),
                        null,
                        "Cell ${crs(rowToCheck, c)} has value ${v(cellToCheck)} when it should not"
                    )
                )
                else -> if (cellToCheck.getValue() != cellToCompare.getValue()) {
                    differences.add(
                        CellContentDifference(
                            cr(rowToCheck, c),
                            cellToCheck.getValue(),
                            cellToCompare.getValue(),
                            "Cell ${crs(
                                rowToCheck,
                                c
                            )} has value ${v(cellToCheck)} when it should be ${v(cellToCompare)}"
                        )
                    )
                }
            }
        }
    }
    return differences
}

sealed class ComparisonResult
object Same : ComparisonResult()
data class Different(val differences: List<Difference>) : ComparisonResult()

sealed class Difference() {
    abstract val message: String
}

data class StructuralDifference(val sheetName: String, override val message: String) : Difference()
data class CellContentDifference(
    val cell: CellReference,
    val left: Any?,
    val right: Any?,
    override val message: String
) :
    Difference()

private fun cr(row: Row, col: Int) = CellReference(row.sheet.sheetName, row.rowNum, col, false, false)
private fun crs(row: Row, col: Int) = cr(row, col).formatAsString()
private fun v(cell: Cell) = cell.getValue()
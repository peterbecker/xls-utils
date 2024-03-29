package de.peterbecker.xls.diff

import de.peterbecker.xls.getValue
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellReference

fun compareWorkbooks(toCheck: Workbook, compareTo: Workbook, diffMode: DiffMode = DiffMode.Strict): ComparisonResult {
    val differences = ArrayList<Difference>()
    toCheck.sheetIterator().forEach {
        when (val compareToSheet = compareTo.getSheet(it.sheetName)) {
            null ->
                if (!diffMode.allowExtraSheets) {
                    differences.add(
                        StructuralDifference(
                            it.sheetName,
                            "Extra sheet present: '${it.sheetName}'"
                        )
                    )
                }
            else -> {
                when (val sheetCompare = compareSheets(it, compareToSheet)) {
                    is Different -> differences.addAll(sheetCompare.differences)
                    is Same -> Unit
                }
            }
        }
    }
    if (!diffMode.allowSheetsMissing) {
        compareTo.sheetIterator().forEach {
            if (toCheck.getSheet(it.sheetName) == null) {
                differences.add(
                    StructuralDifference(
                        it.sheetName,
                        "Sheet missing: '${it.sheetName}'"
                    )
                )
            }
        }
    }
    return if (differences.isEmpty()) Same else Different(differences)
}

fun validateSame(toCheck: Workbook, compareTo: Workbook, diffMode: DiffMode = DiffMode.Strict) {
    when (val result = compareWorkbooks(toCheck, compareTo, diffMode)) {
        is Different -> throw DocumentsDifferException(result.differences)
        is Same -> Unit
    }
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
    if (toCheck.lastRowNum != compareTo.lastRowNum) {
        differences.add(
            StructuralDifference(
                toCheck.sheetName,
                "'${toCheck.sheetName}' has ${toCheck.lastRowNum + 1} rows, we expect ${compareTo.lastRowNum + 1}"
            )
        )
    } else {
        for (r in 0..toCheck.lastRowNum) {
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
    }
    return if (differences.isEmpty()) Same else Different(differences)
}

fun compareRows(rowToCheck: Row, rowToCompare: Row): Collection<Difference> {
    val differences = ArrayList<Difference>()
    if(rowToCheck.firstCellNum != rowToCompare.firstCellNum) {
        differences.add(
            StructuralDifference(
                rowToCheck.sheet.sheetName,
            "Row starts at ${crs(rowToCheck, rowToCheck.firstCellNum.toInt())}, " +
                    "when it should start at ${crs(rowToCompare, rowToCompare.firstCellNum.toInt())}")
        )
    }
    if(rowToCheck.lastCellNum != rowToCompare.lastCellNum) {
        differences.add(
            StructuralDifference(
                rowToCheck.sheet.sheetName,
            "Row ends at ${crs(rowToCheck, rowToCheck.lastCellNum.toInt())}, " +
                    "when it should end at ${crs(rowToCompare, rowToCompare.lastCellNum.toInt())}")
        )
    }
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
                        "Cell ${crs(rowToCheck, c)} has value ${v(cellToCheck)} when it should not have any"
                    )
                )
                else -> if (cellToCheck.getValue() != cellToCompare.getValue()) {
                    differences.add(
                        CellContentDifference(
                            cr(rowToCheck, c),
                            cellToCheck.getValue(),
                            cellToCompare.getValue(),
                            "Cell ${crs(rowToCheck, c)} " +
                                    "has value ${v(cellToCheck)} when it should be ${v(cellToCompare)}"
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

sealed class Difference {
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

class DocumentsDifferException(
    val differences: List<Difference>
) : Exception(
    "Documents differ:\n" + differences.joinToString(prefix = " - ", separator = ",\n") { it.message }
)


private fun cr(row: Row, col: Int) = CellReference(row.sheet.sheetName, row.rowNum, col, false, false)
private fun crs(row: Row, col: Int) = cr(row, col).formatAsString()
private fun v(cell: Cell) = cell.getValue()
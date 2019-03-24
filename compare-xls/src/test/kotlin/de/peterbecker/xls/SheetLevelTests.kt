package de.peterbecker.xls

import org.assertj.core.api.Assertions.assertThat
import org.junit.Test

class SheetLevelTests {
    @Test
    internal fun sameSheetsIsSame() {
        validateSame("sheet_1_2", "sheet_1_2")
    }

    @Test
    internal fun sameSheetsReorderedIsSame() {
        validateSame("sheet_2_1", "sheet_1_2")
    }

    @Test
    internal fun oneSheetDiffers() {
        validateDifferences(
            "sheet_1_2", "sheet_2_3",
            StructuralDifference("Sheet1", "Extra sheet present: 'Sheet1'"),
            StructuralDifference("Sheet3", "Sheet missing: 'Sheet3'")
        )
    }

    @Test
    internal fun oneSheetDiffersVar2() {
        validateDifferences(
            "sheet_2_3", "sheet_1_2",
            StructuralDifference("Sheet3", "Extra sheet present: 'Sheet3'"),
            StructuralDifference("Sheet1", "Sheet missing: 'Sheet1'")
        )
    }

    @Test
    internal fun oneSheetDiffersLenient() {
        validateSame(
            compareWorkbooks(
                wb("sheet_2_3"), wb("sheet_1_2"),
                DiffMode(
                    allowExtraSheets = true,
                    allowSheetsMissing = true
                )
            )
        )
    }

    @Test
    internal fun oneSheetDiffersAllowExtraSheet() {
        validateDifferences(
            compareWorkbooks(
                wb("sheet_2_3"), wb("sheet_1_2"),
                DiffMode(
                    allowExtraSheets = true,
                    allowSheetsMissing = false
                )
            ),
            StructuralDifference("Sheet1", "Sheet missing: 'Sheet1'")
        )
    }

    @Test
    internal fun oneSheetDiffersAllowMissingSheet() {
        validateDifferences(
            compareWorkbooks(
                wb("sheet_2_3"), wb("sheet_1_2"),
                DiffMode(
                    allowExtraSheets = false,
                    allowSheetsMissing = true
                )
            ),
            StructuralDifference("Sheet3", "Extra sheet present: 'Sheet3'")
        )
    }

    @Test
    fun sheetNames() {
        val wb = wb("sheet_1_2")
        val s1 = wb.getSheet("Sheet1")
        val s2 = wb.getSheet("Sheet2")
        assertThat(compareSheets(s1, s1)).isEqualTo(Same)
        validateDifferences(
            compareSheets(s1, s2),
            StructuralDifference("Sheet1", "Sheet names differ: 'Sheet1' instead of 'Sheet2'")
        )
        validateSame(
            compareSheets(s1, s2, DiffMode(allowSheetNameDifference = true))
        )
    }
}
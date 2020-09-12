package de.peterbecker.xls.diff

import org.apache.poi.ss.util.CellReference
import org.junit.jupiter.api.Test

class ValueTests {
    @Test
    fun sameFile() {
        validateSame("3x3matrix_inc", "3x3matrix_inc")
    }

    @Test
    fun oneNumberDiffers() {
        validateDifferences(
            "3x3matrix_inc_center_7", "3x3matrix_inc",
            CellContentDifference(
                CellReference("Sheet1!B2"),
                7.0,
                5.0,
                "Cell Sheet1!B2 has value 7.0 when it should be 5.0"
            )
        )
    }

    @Test
    fun oneSheetHasMoreRows() {
        validateDifferences(
            "3x3matrix_inc", "3x3matrix_inc_extra_row",
            StructuralDifference("Sheet1", "'Sheet1' has 3 rows, we expect 6")
        )
    }

    @Test
    fun oneSheetHasLessRows() {
        validateDifferences(
            "3x3matrix_inc_extra_row", "3x3matrix_inc",
            StructuralDifference("Sheet1", "'Sheet1' has 6 rows, we expect 3")
        )
    }

    // maybe we need to test nulls on either side (but not both)
}
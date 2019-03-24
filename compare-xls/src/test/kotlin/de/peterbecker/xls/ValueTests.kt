package de.peterbecker.xls

import org.apache.poi.ss.util.CellReference
import org.junit.Test

class ValueTests {
    @Test
    fun sameFile() {
        validateSame("3x3matrix_inc", "3x3matrix_inc")
    }

    @Test
    fun oneNumberDiffers() {
        validateDifferences(
            "3x3matrix_inc_center_7", "3x3matrix_inc",
            CellContentDifference(CellReference("Sheet1!B2"), 7.0, 5.0, "Cell Sheet1!B2 has value 7.0 when it should be 5.0")
        )
    }

    // missing tests: more rows, less rows, maybe we need nulls on either side (but not both)
}
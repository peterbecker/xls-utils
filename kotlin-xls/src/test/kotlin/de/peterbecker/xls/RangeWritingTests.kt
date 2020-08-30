package de.peterbecker.xls

import de.peterbecker.xls.diff.validateSame
import org.apache.poi.ss.usermodel.Workbook
import org.assertj.core.api.Assertions.assertThat
import org.junit.Test

class RangeWritingTests {
    @Test
    fun testEmptyData() {
        val act = writeToTestRange(listOf())
        validateSame(act, getTestWorkbook())
    }

    @Test
    fun testEmptyRows() {
        val rows = sequence<List<Any>> { listOf<Any>() }.take(100).toList()
        val act = writeToTestRange(rows)
        validateSame(act, getTestWorkbook())
    }

    @Test(expected = RowTooLongException::class)
    fun testRowTooLong() {
        writeToTestRange(
            listOf(
                listOf(1, 2, 3, 4),
                listOf(),
                listOf(1, 2, 3, 4, 5),
                listOf(1, 2, 3, 4)
            )
        )
    }

    @Test(expected = NamedRangeNotFound::class)
    fun testInvalidName() {
        val wb = getTestWorkbook()
        wb.writeToRange("not_target_range", listOf())
    }

    @Test
    fun testMixedData() {
        val act = writeToTestRange(
            listOf(
                listOf(1,2,3,4),
                listOf("one", "two", "three", "four"),
                listOf(null, null, "justThree", null),
                listOf("one", 2, "three", 4.0)
            )
        )
        validateSame(act, wb("ranges/mixedData"))
    }

    private fun writeToTestRange(rows: Iterable<List<Any?>>): Workbook {
        val wb = getTestWorkbook()
        wb.writeToRange("target_range", rows)
        return wb
    }

    private fun getTestWorkbook() = wb("ranges/namedRange")
}
package de.peterbecker.xls

import de.peterbecker.xls.diff.validateSame
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Test

abstract class RowWritingTests(private val templateFile: String) {
    @Test
    fun testEmptyData() {
        val act = writeToTarget(listOf())
        validateSame(act, getTestWorkbook())
    }

    @Test
    fun testEmptyRows() {
        val rows = sequence<List<Any>> { listOf<Any>() }.take(100).toList()
        val act = writeToTarget(rows)
        validateSame(act, getTestWorkbook())
    }

    @Test(expected = RowTooLongException::class)
    fun testRowTooLong() {
        writeToTarget(
            listOf(
                listOf(1, 2, 3, 4),
                listOf(),
                listOf(1, 2, 3, 4, 5),
                listOf(1, 2, 3, 4)
            )
        )
    }

    @Test(expected = NameNotFound::class)
    fun testInvalidName() {
        val wb = getTestWorkbook()
        wb.writeToRange("not_target_range", listOf())
    }

    @Test
    fun testMixedData() {
        val act = writeToTarget(
            listOf(
                listOf(1,2,3,4),
                listOf("one", "two", "three", "four"),
                listOf(null, null, "justThree", null),
                listOf("one", 2, "three", 4.0)
            )
        )
        validateSame(act, wb("ranges/mixedData"))
    }

    private fun writeToTarget(rows: Iterable<List<Any?>>): Workbook {
        val wb = getTestWorkbook()
        writeToTarget(wb, "target_range", rows)
        return wb
    }

    private fun getTestWorkbook() = wb(templateFile)

    protected abstract fun writeToTarget(wb: XSSFWorkbook, name: String, rows: Iterable<List<Any?>>)
}
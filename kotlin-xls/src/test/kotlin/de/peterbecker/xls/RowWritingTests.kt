package de.peterbecker.xls

import de.peterbecker.xls.diff.validateSame
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Assertions
import org.junit.jupiter.api.Test

abstract class RowWritingTests(private val templateFile: String, private val type: String) {
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

    @Test
    fun testRowTooLong() {
        val exception = Assertions.assertThrows(RowTooLongException::class.java) {
            writeToTarget(
                listOf(
                    listOf(1, 2, 3, 4),
                    listOf(),
                    listOf(1, 2, 3, 4, 5),
                    listOf(1, 2, 3, 4)
                )
            )
        }
        assertThat(exception.rowNumber).isEqualTo(2)
        assertThat(exception.rangeWidth).isEqualTo(4)
        assertThat(exception.rowLength).isEqualTo(5)
    }

    @Test
    fun testInvalidName() {
        val wb = getTestWorkbook()
        val exception = Assertions.assertThrows(NameNotFound::class.java) {
            writeToTarget(wb, "not_target_range", listOf())
        }
        assertThat(exception.name).isEqualTo("not_target_range")
        assertThat(exception.type).isEqualTo(type)
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
package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet
import org.assertj.core.api.Assertions.assertThat
import org.assertj.core.data.Offset.offset
import org.junit.jupiter.api.Test

class CellValueTests {
    @Test
    fun `read basic values`() {
        val workbook = wb("values")
        val sheet = workbook.getSheet("basic")
        assertThat(sheet.getDoubleValueAt(0, 0)).isCloseTo(1.0, offset(0.000001))
        assertThat(sheet.getValueAt(0, 1)).isEqualTo(true)
        assertThat(sheet.getValueAt(0, 2)).isEqualTo("Hello World")
        assertThat(sheet.getDoubleValueAt(0, 3)).isCloseTo(5.67, offset(0.000001))
    }

    private fun Sheet.getDoubleValueAt(row: Int, col: Int) = this.getValueAt(row, col) as Double?

    @Test
    fun `read absent cells`() {
        val workbook = wb("values")
        val sheet = workbook.getSheet("basic")
        assertThat(sheet.getValueAt(100, 100)).isNull()
    }
}
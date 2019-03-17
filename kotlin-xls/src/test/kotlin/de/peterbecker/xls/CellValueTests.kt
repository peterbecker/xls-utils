package de.peterbecker.xls

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.assertj.core.api.Assertions.assertThat
import org.assertj.core.data.Offset
import org.junit.Test

class CellValueTests {
    @Test
    fun getBasicValues() {
        val workbook = loadWorkbook("values")
        val sheet = workbook.getSheet("basic")
        assertThat(sheet.getDoubleValueAt(0, 0)).isCloseTo(1.0, Offset.offset(0.000001))
        assertThat(sheet.getValueAt(0, 1)).isEqualTo(true)
        assertThat(sheet.getValueAt(0, 2)).isEqualTo("Hello World")
        assertThat(sheet.getDoubleValueAt(0, 3)).isCloseTo(5.67, Offset.offset(0.000001))
    }

    private fun loadWorkbook(name: String) =
        WorkbookFactory.create(CellValueTests::class.java.getResourceAsStream("/${name}.xlsx"))
}
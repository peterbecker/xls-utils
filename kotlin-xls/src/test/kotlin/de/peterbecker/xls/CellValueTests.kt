package de.peterbecker.xls

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.assertj.core.api.Assertions.assertThat
import org.assertj.core.data.Offset.offset
import org.junit.Assert.fail
import org.junit.Test

class CellValueTests {
    @Test
    fun getBasicValues() {
        val workbook = loadWorkbook("values")
        val sheet = workbook.getSheet("basic")
        assertThat(sheet.getDoubleValueAt(0, 0)).isCloseTo(1.0, offset(0.000001))
        assertThat(sheet.getValueAt(0, 1)).isEqualTo(true)
        assertThat(sheet.getBooleanValueAt(0, 1)).isEqualTo(true)
        assertThat(sheet.getValueAt(0, 2)).isEqualTo("Hello World")
        assertThat(sheet.getStringValueAt(0, 2)).isEqualTo("Hello World")
        assertThat(sheet.getDoubleValueAt(0, 3)).isCloseTo(5.67, offset(0.000001))
    }

    @Test
    fun testMissingValues() {
        val workbook = loadWorkbook("values")
        val sheet = workbook.getSheet("basic")
        assertThat(sheet.getValueAt(100, 100)).isNull()
        assertThat(sheet.getBooleanValueAt(100, 100)).isNull()
        assertThat(sheet.getDoubleValueAt(100, 100)).isNull()
        assertThat(sheet.getStringValueAt(100, 100)).isNull()
    }

    @Test
    fun testTypeMismatches() {
        val workbook = loadWorkbook("values")
        val sheet = workbook.getSheet("basic")
        try {
            sheet.getBooleanValueAt(0, 0)
            fail("Should have thrown exception")
        } catch (e: ClassCastException) {
            // expected
        }
        try {
            sheet.getStringValueAt(0, 0)
            fail("Should have thrown exception")
        } catch (e: ClassCastException) {
            // expected
        }
        try {
            sheet.getDoubleValueAt(0, 1)
            fail("Should have thrown exception")
        } catch (e: ClassCastException) {
            // expected
        }
        try {
            sheet.getStringValueAt(0, 1)
            fail("Should have thrown exception")
        } catch (e: ClassCastException) {
            // expected
        }
        try {
            sheet.getBooleanValueAt(0, 2)
            fail("Should have thrown exception")
        } catch (e: ClassCastException) {
            // expected
        }
        try {
            sheet.getDoubleValueAt(0, 2)
            fail("Should have thrown exception")
        } catch (e: ClassCastException) {
            // expected
        }
    }

    private fun loadWorkbook(name: String) =
        WorkbookFactory.create(CellValueTests::class.java.getResourceAsStream("/${name}.xlsx"))
}
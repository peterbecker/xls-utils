package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFTable
import org.assertj.core.api.Assertions.assertThat
import org.assertj.core.data.Offset
import org.junit.Test

class LoadTablesTests : TestBase() {
    @Test
    fun `test load two tables`() {
        val act = con().use {
            runReports(wb("countries_template_tables"), it)
        }
        val sheet = act.getSheet("Results") as XSSFSheet
        assertThat(sheet.getValueAt(2, 0)).isEqualTo("China")
        assertThat(sheet.getDoubleValueAt(2, 1)).isCloseTo(1392730000.0, Offset.offset(0.000001))
        assertThat(sheet.getValueAt(11, 0)).isEqualTo("Japan")
        assertThat(sheet.getDoubleValueAt(11, 1)).isCloseTo(126529100.0, Offset.offset(0.000001))
        assertThat(sheet.getValueAt(2, 3)).isEqualTo("Russia")
        assertThat(sheet.getDoubleValueAt(2, 4)).isCloseTo(17100000.0, Offset.offset(0.000001))
        assertThat(sheet.getValueAt(11, 3)).isEqualTo("Kazakhstan")
        assertThat(sheet.getDoubleValueAt(11, 4)).isCloseTo(2717300.0, Offset.offset(0.000001))
        // should have dropped the query sheets
        assertThat(act.numberOfSheets).isEqualTo(1)
        // ranges should match the new areas
        assertThat(sheet.tables.first { it.name == "Population" }).matches(
            { isTable(it, 1, 0, 11, 1) },
            "Table area matches"
        )
        assertThat(sheet.tables.first { it.name == "Size" }).matches(
            { isTable(it, 1, 3, 11, 4) },
            "Table area matches"
        )
    }

    private fun isTable(table: XSSFTable, top: Int, left: Short, bottom: Int, right: Short): Boolean {
        return table.area.firstCell.row == top &&
                table.area.firstCell.col == left &&
                table.area.lastCell.row == bottom &&
                table.area.lastCell.col == right
    }
}
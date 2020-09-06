package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.usermodel.Name
import org.apache.poi.ss.util.AreaReference
import org.assertj.core.api.Assertions.assertThat
import org.assertj.core.data.Offset
import org.junit.jupiter.api.Test

class LoadRangesTests : TestBase() {
    @Test
    fun `test load two ranges`() {
        val act = con().use{
            runReports(wb("countries_template_ranges"), it)
        }
        val popSheet = act.getSheet("Countries by Population")
        assertThat(popSheet.getValueAt(1, 0)).isEqualTo("China")
        assertThat(popSheet.getDoubleValueAt(1, 1)).isCloseTo(1392730000.0, Offset.offset(0.000001))
        assertThat(popSheet.getValueAt(10, 0)).isEqualTo("Japan")
        assertThat(popSheet.getDoubleValueAt(10, 1)).isCloseTo(126529100.0, Offset.offset(0.000001))
        val sizeSheet = act.getSheet("Countries by Size")
        assertThat(sizeSheet.getValueAt(1, 0)).isEqualTo("Russia")
        assertThat(sizeSheet.getDoubleValueAt(1, 1)).isCloseTo(17100000.0, Offset.offset(0.000001))
        assertThat(sizeSheet.getValueAt(10, 0)).isEqualTo("Kazakhstan")
        assertThat(sizeSheet.getDoubleValueAt(10, 1)).isCloseTo(2717300.0, Offset.offset(0.000001))
        // should have dropped the query sheets
        assertThat(act.numberOfSheets).isEqualTo(2)
        // ranges should match the new areas
        assertThat(act.getName("Population")).matches(
            { isRange(it, "Countries by Population", 1, 0, 11, 1) },
            "Range does not match"
        )
        assertThat(act.getName("Size")).matches(
            { isRange(it, "Countries by Size", 1, 0, 11, 1) },
            "Range does not match"
        )
    }

    private fun isRange(name: Name, sheetName: String, top: Int, left: Short, bottom: Int, right: Short): Boolean {
        val area = AreaReference(name.refersToFormula, SpreadsheetVersion.EXCEL2007)
        return name.sheetName == sheetName &&
                area.firstCell.row == top &&
                area.firstCell.col == left &&
                area.lastCell.row == bottom &&
                area.lastCell.col == right
    }
}
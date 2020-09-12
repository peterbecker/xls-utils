package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFChart
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test

class ChartTests {
    @Test
    internal fun `find chart on sheet by title`() {
        val workbook = wb("charts")
        val sheet = workbook.getSheet("Sheet1")
        assertThat(sheet.findChartByTitle("Test Title")).isInstanceOf(XSSFChart::class.java)
        assertThat(sheet.findChartByTitle("Not a valid title")).isNull()
    }

    @Test
    internal fun `find charts in workbook by title`() {
        val workbook = wb("charts")
        assertThat(workbook.findChartByTitle("Test Title")).isInstanceOf(XSSFChart::class.java)
        assertThat(workbook.findChartByTitle("Chart on Separate Sheet")).isInstanceOf(XSSFChart::class.java)
        assertThat(workbook.findChartByTitle("Not a valid title")).isNull()
    }
}

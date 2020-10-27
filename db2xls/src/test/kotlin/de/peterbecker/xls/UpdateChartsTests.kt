package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.xssf.usermodel.XSSFChart
import org.junit.jupiter.api.Disabled
import org.junit.jupiter.api.Test

@Disabled("Not yet implemented")
// see https://stackoverflow.com/questions/56159322/change-data-range-in-excel-line-chart-using-apache-poi
class UpdateChartsTests : TestBase() {
    @Test
    fun `validate chart references update`() {
        val act = con().use {
            runReports(wb("countries_template_charts"), it)
        }
        // for charts it's complicated as it's specific to chart type and split across series (labels) and values
        val popChart = act.findChartByTitle("Population")!! // chart with a specific series don't seem to have the same kind of title, this misses at the moment
        assertThat(getPieChartSeriesRef(popChart)).hasRange(2, 0, 11, 0)
        assertThat(getPieChartValuesRef(popChart)).hasRange(2, 1, 11, 1)
        val sizeChart = act.findChartByTitle("Size")!!
        assertThat(getPieChartSeriesRef(sizeChart)).hasRange(2, 0, 11, 0)
        assertThat(getPieChartValuesRef(sizeChart)).hasRange(2, 2, 11, 2)
        val donutChart = act.findChartByTitle("Size outside Population")!!
        assertThat(getPieChartSeriesRef(donutChart)).hasRange(2, 0, 11, 0)
        assertThat(getPieChartValuesRef(donutChart)).hasRange(2, 1, 11, 1)
        assertThat(getPieChartSeriesRef(donutChart,1)).hasRange(2, 0, 11, 0)
        assertThat(getPieChartValuesRef(donutChart,1)).hasRange(2, 2, 11, 2)
        val lineChart = act.findChartByTitle("Lines with Two Axes")!!
        assertThat(getLineChartSeriesRef(lineChart)).hasRange(2, 0, 11, 0)
        assertThat(getLineChartValuesRef(lineChart)).hasRange(2, 1, 11, 1)
        assertThat(getLineChartSeriesRef(lineChart,1)).hasRange(2, 0, 11, 0)
        assertThat(getLineChartValuesRef(lineChart,1)).hasRange(2, 2, 11, 2)
    }

    private fun getPieChartSeriesRef(chart: XSSFChart, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getPieChartArray(0).getSerArray(series).cat.strRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getPieChartValuesRef(chart: XSSFChart, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getPieChartArray(0).getSerArray(series).`val`.numRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getLineChartSeriesRef(chart: XSSFChart, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getLineChartArray(0).getSerArray(series).cat.strRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getLineChartValuesRef(chart: XSSFChart, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getLineChartArray(0).getSerArray(series).`val`.numRef.f, SpreadsheetVersion.EXCEL2007)
}
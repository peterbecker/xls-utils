package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.xssf.usermodel.XSSFChart
import org.junit.jupiter.api.Disabled
import org.junit.jupiter.api.Test
import java.io.FileOutputStream

class UpdateChartsTests : TestBase() {
    @Test
    fun `validate chart references update`() {
        val act = con().use {
            runReports(wb("countries_template_charts"), it)
        }
        debugWrite(act, "validateChartReferencesUpdate")
        // for charts it's complicated as it's specific to chart type and split across series (labels) and values
        val popChart = act.findChartByTitle("Population")!! // chart with a specific series don't seem to have the same kind of title, this misses at the moment
        assertThat(getPieChartSeriesRef(popChart)).hasRange(2, 0, 11, 0)
        assertThat(getPieChartValuesRef(popChart)).hasRange(2, 1, 11, 1)
        val sizeChart = act.findChartByTitle("Size")!!
        assertThat(getPie3DChartSeriesRef(sizeChart)).hasRange(2, 0, 11, 0)
        assertThat(getPie3DChartValuesRef(sizeChart)).hasRange(2, 2, 11, 2)
        val donutChart = act.findChartByTitle("Size outside Population")!!
        assertThat(getDoughnutChartSeriesRef(donutChart)).hasRange(2, 0, 11, 0)
        assertThat(getDoughnutChartValuesRef(donutChart)).hasRange(2, 1, 11, 1)
        assertThat(getDoughnutChartSeriesRef(donutChart,series = 1)).hasRange(2, 0, 11, 0)
        assertThat(getDoughnutChartValuesRef(donutChart,series = 1)).hasRange(2, 2, 11, 2)
        val lineChart = act.findChartByTitle("Lines with Two Axes")!!
        assertThat(getLineChartSeriesRef(lineChart)).hasRange(2, 0, 11, 0)
        assertThat(getLineChartValuesRef(lineChart)).hasRange(2, 1, 11, 1)
        assertThat(getLineChartSeriesRef(lineChart,1)).hasRange(2, 0, 11, 0)
        assertThat(getLineChartValuesRef(lineChart,1)).hasRange(2, 2, 11, 2)
    }

    private fun getPieChartSeriesRef(chart: XSSFChart, ind: Int = 0, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getPieChartArray(ind).getSerArray(series).cat.strRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getPieChartValuesRef(chart: XSSFChart, ind: Int = 0, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getPieChartArray(ind).getSerArray(series).`val`.numRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getPie3DChartSeriesRef(chart: XSSFChart, ind: Int = 0, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getPie3DChartArray(ind).getSerArray(series).cat.strRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getPie3DChartValuesRef(chart: XSSFChart, ind: Int = 0, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getPie3DChartArray(ind).getSerArray(series).`val`.numRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getDoughnutChartSeriesRef(chart: XSSFChart, ind: Int = 0, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getDoughnutChartArray(ind).getSerArray(series).cat.strRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getDoughnutChartValuesRef(chart: XSSFChart, ind: Int = 0, series: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getDoughnutChartArray(ind).getSerArray(series).`val`.numRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getLineChartSeriesRef(chart: XSSFChart, ind: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getLineChartArray(ind).getSerArray(0).cat.strRef.f, SpreadsheetVersion.EXCEL2007)

    private fun getLineChartValuesRef(chart: XSSFChart, ind: Int = 0) =
            AreaReference(chart.ctChart.plotArea.getLineChartArray(ind).getSerArray(0).`val`.numRef.f, SpreadsheetVersion.EXCEL2007)
}
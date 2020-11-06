@file:Suppress("DuplicatedCode") // the XML classes don't have common base types and we can't change that

package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.xssf.usermodel.XSSFChart
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef


// Credit goes to https://stackoverflow.com/a/56196387/19820

/**
 * @see [XSSFWorkbook.expandChartReferences]
 */
fun XSSFChart.expandReferences(area: AreaReference) {
    val plotArea = this.ctChart.plotArea
    plotArea.pieChartList.forEach { chart ->
        chart.serList.forEach { series ->
            updateStrRef(series.cat.strRef, area)
            updateNumRef(series.`val`.numRef, area)
        }
    }
    plotArea.pie3DChartList.forEach { chart ->
        chart.serList.forEach { series ->
            updateStrRef(series.cat.strRef, area)
            updateNumRef(series.`val`.numRef, area)
        }
    }
    plotArea.doughnutChartList.forEach { chart ->
        chart.serList.forEach { series ->
            updateStrRef(series.cat.strRef, area)
            updateNumRef(series.`val`.numRef, area)
        }
    }
    plotArea.lineChartList.forEach { chart ->
        chart.serList.forEach { series ->
            updateStrRef(series.cat.strRef, area)
            updateNumRef(series.`val`.numRef, area)
        }
    }
}

private fun updateStrRef(ctRef: CTStrRef, area: AreaReference) {
    val ref = AreaReference(ctRef.f, SpreadsheetVersion.EXCEL2007)
    if (area.contains(ref, true) && area.firstCell.row == ref.firstCell.row) {
        val newRef = AreaReference(ref.firstCell, ref.lastCell.replace(row = area.lastCell.row), SpreadsheetVersion.EXCEL2007)
        ctRef.f = newRef.formatAsString()
    }
}

private fun updateNumRef(ctRef: CTNumRef, area: AreaReference) {
    val ref = AreaReference(ctRef.f, SpreadsheetVersion.EXCEL2007)
    if (area.contains(ref, true) && area.firstCell.row == ref.firstCell.row) {
        val newRef = AreaReference(ref.firstCell, ref.lastCell.replace(row = area.lastCell.row), SpreadsheetVersion.EXCEL2007)
        ctRef.f = newRef.formatAsString()
    }
}

package de.peterbecker.xls

import org.apache.poi.ss.util.AreaReference
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook


fun XSSFSheet.findChartByTitle(title: String) =
        this.drawingPatriarch.charts.find { it.titleText.string == title }

/**
 * @see [XSSFWorkbook.expandChartReferences]
 */
fun XSSFSheet.expandChartReferences(area: AreaReference) {
    this.drawingPatriarch?.charts?.forEach {
        it.expandReferences(area)
    }
}
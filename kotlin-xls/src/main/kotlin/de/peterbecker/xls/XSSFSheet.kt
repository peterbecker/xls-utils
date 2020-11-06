package de.peterbecker.xls

import org.apache.poi.ss.util.AreaReference
import org.apache.poi.xssf.usermodel.XSSFChart
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.w3c.dom.CharacterData
import org.w3c.dom.Node


fun XSSFSheet.findChartByTitle(title: String) =
        this.drawingPatriarch.charts.find { it.titleText.string == title || hasTitleInside(it, title) }

/*
Pie chart titles seem to be just somewhere inside, others might be as well.
 */
private fun hasTitleInside(chart: XSSFChart, title: String) =
        containsText(chart.ctChart.plotArea.domNode, title)

private fun containsText(node: Node, text: String): Boolean {
    when (node) {
        is CharacterData -> if(node.data == text) return true
    }
    for (i in 0 until node.childNodes.length) {
        if (containsText(node.childNodes.item(i), text)) return true
    }
    return false
}


/**
 * @see [XSSFWorkbook.expandChartReferences]
 */
fun XSSFSheet.expandChartReferences(area: AreaReference) {
    this.drawingPatriarch?.charts?.forEach {
        it.expandReferences(area)
    }
}
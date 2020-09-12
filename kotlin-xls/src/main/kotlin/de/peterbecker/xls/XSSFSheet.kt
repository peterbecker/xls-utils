package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFSheet


fun XSSFSheet.findChartByTitle(title: String) =
        this.drawingPatriarch.charts.find { it.titleText.string == title }

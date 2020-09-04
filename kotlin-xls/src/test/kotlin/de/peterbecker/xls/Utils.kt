package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory

fun wb(name: String): XSSFWorkbook {
    Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
        return XSSFWorkbookFactory.create(it) as XSSFWorkbook
    }
}
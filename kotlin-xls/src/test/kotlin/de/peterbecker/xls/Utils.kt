package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

fun wb(name: String): Workbook {
    Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
        return WorkbookFactory.create(it)
    }
}
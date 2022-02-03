package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun wb(name: String): XSSFWorkbook {
    Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
        return XSSFWorkbook(it)
    }
}
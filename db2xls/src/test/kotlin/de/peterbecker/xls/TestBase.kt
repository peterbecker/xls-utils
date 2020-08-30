package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.sql.Connection
import java.sql.DriverManager

abstract class TestBase {
    protected fun con(): Connection = DriverManager.getConnection(
        "jdbc:h2:mem:test;INIT=RUNSCRIPT FROM '${resourceUrl("init.sql")}'",
        "sa",
        ""
    )

    fun wb(name: String): Workbook {
        Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
            return WorkbookFactory.create(it)
        }
    }

    protected fun resourceUrl(file: String) = Thread.currentThread().contextClassLoader.getResource(file)!!
}

fun Sheet.getDoubleValueAt(row: Int, col: Int) = this.getValueAt(row, col) as Double?

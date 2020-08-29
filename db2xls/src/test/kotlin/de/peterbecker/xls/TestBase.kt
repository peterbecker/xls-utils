package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.h2.jdbcx.JdbcDataSource
import javax.sql.DataSource

abstract class TestBase {
    protected fun ds(): DataSource {
        val ds = JdbcDataSource()
        ds.setURL("jdbc:h2:mem:test;INIT=RUNSCRIPT FROM '${resourceUrl("init.sql")}'")
        ds.user = "sa"
        ds.password = ""
        return ds
    }

    fun wb(name: String): Workbook {
        Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
            return WorkbookFactory.create(it)
        }
    }

    protected fun resourceUrl(file: String) = Thread.currentThread().contextClassLoader.getResource(file)!!
}

fun Sheet.getDoubleValueAt(row: Int, col: Int) = this.getValueAt(row, col) as Double?

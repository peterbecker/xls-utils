package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import java.nio.file.Path
import java.sql.Connection
import java.sql.DriverManager

abstract class TestBase {
    protected fun con(): Connection = DriverManager.getConnection(
            "jdbc:h2:mem:test;INIT=RUNSCRIPT FROM '${resourceUrl("init.sql")}'",
            "sa",
            ""
    )

    fun wb(name: String): XSSFWorkbook {
        Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
            return XSSFWorkbook(it)
        }
    }

    protected fun resourceUrl(file: String) = Thread.currentThread().contextClassLoader.getResource(file)!!
}

fun Sheet.getDoubleValueAt(row: Int, col: Int) = this.getValueAt(row, col) as Double?

fun debugWrite(wb: Workbook, name: String) =
        System.getenv("OUTPUT_FOLDER")?.let { wb.write(FileOutputStream(Path.of(it, "$name.xlsx").toFile())) }
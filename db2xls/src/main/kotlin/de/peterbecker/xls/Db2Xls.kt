package de.peterbecker.xls

import com.sksamuel.hoplite.ConfigLoader
import mu.KotlinLogging
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.sql.Connection
import java.sql.DriverManager
import java.sql.ResultSet
import kotlin.system.exitProcess

private val logger = KotlinLogging.logger {}

const val queryPrefix = "Query_"

fun runReports(wb: XSSFWorkbook, con: Connection): XSSFWorkbook {
    val toRemove = mutableListOf<Int>()
    wb.sheetIterator().forEach {
        val name = it.sheetName
        if (name.startsWith(queryPrefix, ignoreCase = true) && name.length > queryPrefix.length) {
            processQuery(wb, it, con)
            toRemove += wb.getSheetIndex(it)
        }
    }
    toRemove.reversed().forEach {
        wb.removeSheetAt(it)
    }
    return wb
}

fun processQuery(wb: XSSFWorkbook, sheet: Sheet, con: Connection) {
    logger.info { "Processing ${sheet.sheetName}" }
    when (val query = sheet.getValueAt(0, 0)) {
        is String -> {
            processQuery(wb, query, sheet.sheetName.substring(queryPrefix.length), con)
        }
        else -> {
            throw InvalidInputException("No query in top left cell of ${sheet.sheetName}")
        }
    }
}

fun processQuery(wb: XSSFWorkbook, query: String, targetName: String, con: Connection) {
    con.createStatement().use { stmt ->
        stmt.executeQuery(query).use { rs ->
            val area = wb.writeData(targetName, toLists(rs))
            wb.expandChartReferences(area)
        }
    }
}

fun toLists(rs: ResultSet): Iterable<List<Any?>> = rs.use {
    generateSequence {
        if (rs.next()) {
            (1..rs.metaData.columnCount).map {
                rs.getObject(it)
            }.toList()
        } else null
    }.toList()
}

fun main(args: Array<String>) {
    when (args.size) {
        1 -> {
            val config = ConfigLoader().loadConfigOrThrow<Config>(args[0])
            val template = XSSFWorkbook(File(config.template))
            val result = DriverManager.getConnection(config.db.url, config.db.user, config.db.password).use {
                runReports(template, it)
            }
            result.write(FileOutputStream(config.output))
        }
        else -> {
            println("Please provide a path to a configuration file as sole argument.")
            exitProcess(1)
        }
    }
}

class InvalidInputException(override val message: String) : Exception(message)

data class DatabaseConfig(
    val url: String,
    val user: String,
    val password: String
)

data class Config(
    val db: DatabaseConfig,
    val template: String,
    val output: String
)
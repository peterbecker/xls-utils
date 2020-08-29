package de.peterbecker.xls

import mu.KotlinLogging
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import java.sql.Connection
import java.sql.ResultSet
import javax.sql.DataSource

private val logger = KotlinLogging.logger {}

const val queryPrefix = "Query_"

fun runReports(wb: Workbook, data: DataSource): Workbook {
    data.connection.use {con ->
        val toRemove = mutableListOf<Int>()
        wb.sheetIterator().forEach {
            val name = it.sheetName
            if(name.startsWith(queryPrefix, ignoreCase = true) && name.length > queryPrefix.length) {
                processQuery(wb, it, con)
                toRemove += wb.getSheetIndex(it)
            }
        }
        toRemove.reversed().forEach {
            wb.removeSheetAt(it)
        }
    }
    return wb
}

fun processQuery(wb: Workbook, sheet: Sheet, con: Connection) {
    logger.info { "Processing ${sheet.sheetName}" }
    when (val query = sheet.getValueAt(0,0)) {
        is String -> {
            processQuery(wb, query, sheet.sheetName.substring(queryPrefix.length), con)
        }
        else -> {
            throw InvalidInputException("No query in top left cell of ${sheet.sheetName}")
        }
    }
}

fun processQuery(wb: Workbook, query: String, targetName: String, con: Connection) {
    con.createStatement().use { stmt ->
        stmt.executeQuery(query).use { rs ->
            wb.writeToRange(targetName, toLists(rs))
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

class InvalidInputException(override val message: String) : Exception(message)

package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFWorkbook

class TableWritingTests: RowWritingTests("ranges/namedTable", "table") {
    override fun writeToTarget(wb: XSSFWorkbook, name: String, rows: Iterable<List<Any?>>) = wb.writeToTable(name, rows)
}
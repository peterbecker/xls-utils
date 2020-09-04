package de.peterbecker.xls

import org.apache.poi.xssf.usermodel.XSSFWorkbook

class RangeWritingTests : RowWritingTests("ranges/namedRange") {
    override fun writeToTarget(wb: XSSFWorkbook, name: String, rows: Iterable<List<Any?>>) = wb.writeToRange(name, rows)
}

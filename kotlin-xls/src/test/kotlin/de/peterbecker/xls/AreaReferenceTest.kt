package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test

internal class AreaReferenceTest {
    @Test
    internal fun `area contains other`() {
        val area = AreaReference("SheetX!C3:G5", SpreadsheetVersion.EXCEL2007)
        assertThat(area.contains(area)).isTrue
        assertThat(area.contains(AreaReference("SheetY!C3:G5", SpreadsheetVersion.EXCEL2007))).isFalse
        assertThat(area.contains(AreaReference("SheetX!C2:G5", SpreadsheetVersion.EXCEL2007))).isFalse
        assertThat(area.contains(AreaReference("SheetX!B3:G5", SpreadsheetVersion.EXCEL2007))).isFalse
        assertThat(area.contains(AreaReference("SheetX!C3:G6", SpreadsheetVersion.EXCEL2007))).isFalse
        assertThat(area.contains(AreaReference("SheetX!C3:H5", SpreadsheetVersion.EXCEL2007))).isFalse
        assertThat(area.contains(AreaReference("SheetX!H30:AA50", SpreadsheetVersion.EXCEL2007))).isFalse
        assertThat(area.contains(AreaReference("SheetX!C3:G5", SpreadsheetVersion.EXCEL2007))).isTrue
        assertThat(area.contains(AreaReference("SheetX!C4:G5", SpreadsheetVersion.EXCEL2007))).isTrue
        assertThat(area.contains(AreaReference("SheetX!D3:G5", SpreadsheetVersion.EXCEL2007))).isTrue
        assertThat(area.contains(AreaReference("SheetX!C3:G4", SpreadsheetVersion.EXCEL2007))).isTrue
        assertThat(area.contains(AreaReference("SheetX!C3:F5", SpreadsheetVersion.EXCEL2007))).isTrue
        assertThat(area.contains(AreaReference("SheetX!E4:E4", SpreadsheetVersion.EXCEL2007))).isTrue
    }
}
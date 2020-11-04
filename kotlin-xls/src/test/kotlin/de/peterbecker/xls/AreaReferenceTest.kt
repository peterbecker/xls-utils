package de.peterbecker.xls

import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.util.AreaReference
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test

internal class AreaReferenceTest {
    @Test
    internal fun `area contains other`() {
        val area = ar("SheetX!C3:G5")
        assertThat(area.contains(area)).isTrue
        assertThat(area.contains(ar("SheetY!C3:G5"))).isFalse
        assertThat(area.contains(ar("SheetX!C2:G5"))).isFalse
        assertThat(area.contains(ar("SheetX!B3:G5"))).isFalse
        assertThat(area.contains(ar("SheetX!C3:G6"))).isFalse
        assertThat(area.contains(ar("SheetX!C3:H5"))).isFalse
        assertThat(area.contains(ar("SheetX!H30:AA50"))).isFalse
        assertThat(area.contains(ar("SheetX!C3:G5"))).isTrue
        assertThat(area.contains(ar("SheetX!C4:G5"))).isTrue
        assertThat(area.contains(ar("SheetX!D3:G5"))).isTrue
        assertThat(area.contains(ar("SheetX!C3:G4"))).isTrue
        assertThat(area.contains(ar("SheetX!C3:F5"))).isTrue
        assertThat(area.contains(ar("SheetX!E4:E4"))).isTrue
    }

    @Test
    internal fun heights() {
        assertThat(ar("SheetX!C3:C2").height).isEqualTo(0)
        assertThat(ar("SheetX!C3:C3").height).isEqualTo(1)
        assertThat(ar("SheetX!C3:G7").height).isEqualTo(5)
        assertThat(ar("SheetX!A1:A999999").height).isEqualTo(999999)
    }

    @Test
    internal fun widths() {
        assertThat(ar("SheetX!C3:B5").width).isEqualTo(0)
        assertThat(ar("SheetX!C3:C5").width).isEqualTo(1)
        assertThat(ar("SheetX!C3:G5").width).isEqualTo(5)
        assertThat(ar("SheetX!A3:ZZ5").width).isEqualTo(27 * 26)
    }

    private fun ar(ref: String) = AreaReference(ref, SpreadsheetVersion.EXCEL2007)
}
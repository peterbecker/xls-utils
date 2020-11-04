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

    @Test
    internal fun `area adjustments`() {
        val area = ar("SheetX!D3:G7")
        assertThat(area.derive().formatAsString()).isEqualTo("SheetX!D3:G7")
        assertThat(area.derive(top=3).formatAsString()).isEqualTo("SheetX!D6:G7")
        assertThat(area.derive(top=-2).formatAsString()).isEqualTo("SheetX!D1:G7")
        assertThat(area.derive(left=4).formatAsString()).isEqualTo("SheetX!G3:H7") // columns flipped
        assertThat(area.derive(left=-1).formatAsString()).isEqualTo("SheetX!C3:G7")
        assertThat(area.derive(bottom=4).formatAsString()).isEqualTo("SheetX!D3:G11")
        assertThat(area.derive(bottom=-5).formatAsString()).isEqualTo("SheetX!D2:G3") // rows flipped
        assertThat(area.derive(right=3).formatAsString()).isEqualTo("SheetX!D3:J7")
        assertThat(area.derive(right=-2).formatAsString()).isEqualTo("SheetX!D3:E7")
        assertThat(area.derive(top = -1, left = -1, bottom = 1, right = 1).formatAsString()).isEqualTo("SheetX!C2:H8")
    }

    private fun ar(ref: String) = AreaReference(ref, SpreadsheetVersion.EXCEL2007)
}
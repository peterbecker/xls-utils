package de.peterbecker.xls

import org.apache.poi.ss.util.AreaReference
import org.assertj.core.api.AbstractAssert

class AreaReferenceAssert(areaReference: AreaReference):
        AbstractAssert<AreaReferenceAssert, AreaReference>(areaReference, AreaReferenceAssert::class.java) {
    fun hasRange(top: Int, left: Short, bottom: Int, right: Short) {
        isNotNull
        if (actual.firstCell.row != top ||
                actual.firstCell.col != left ||
                actual.lastCell.row != bottom ||
                actual.lastCell.col != right) {
            failWithMessage(
                    "Expected range to be from (${top},${left}) to (${bottom},${right}), but got " +
                            "(${actual.firstCell.row},${actual.firstCell.col}) to (${actual.lastCell.row},${ actual.lastCell.col})"
            )
        }
    }
}

fun assertThat(areaReference: AreaReference) = AreaReferenceAssert(areaReference)
package de.peterbecker.xls

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.assertj.core.api.Assertions.assertThat

fun wb(name: String): Workbook {
    Thread.currentThread().contextClassLoader.getResourceAsStream("$name.xlsx").use {
        return WorkbookFactory.create(it)
    }
}

fun validateSame(left: String, right: String) {
    val result = compareWorkbooks(wb(left), wb(right))
    validateSame(result)
}

fun validateSame(result: ComparisonResult) {
    when (result) {
        is Different ->
            throw AssertionError(
                "Expected workbooks to be same, but found differences:\n" +
                        result.differences.joinToString("\n") { it.message }
            )
    }
}

fun validateDifferences(left: String, right: String, vararg expected: Difference) =
        validateDifferences(compareWorkbooks(wb(left), wb(right)), *expected)

fun validateDifferences(result: ComparisonResult, vararg expected: Difference) {
    when(result) {
        is Different -> assertThat(result.differences).containsExactlyInAnyOrder(*expected)
        else -> throw AssertionError("Expected differences, but found none")
    }
}
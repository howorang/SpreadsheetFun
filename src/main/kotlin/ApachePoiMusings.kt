import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import java.math.BigDecimal
import java.time.LocalDateTime


fun setUpWorkBook(): XSSFWorkbook {
    val workBook = XSSFWorkbook()
    val sheet = workBook.createSheet()

    val kdbData = getKdbData()
    val columns = getKdbColumns()

    val headerRow = sheet.createRow(0)
    for ((cellIndex, cellValue) in columns.withIndex()) {
        val cell = headerRow.createCell(cellIndex)
        cell.setCellValue(cellValue)
    }

    for ((rowIndex, row) in kdbData.withIndex()) {
        val currentRow = sheet.createRow(rowIndex + 1)
        for ((cellIndex, value) in row.values.withIndex()) {
            val cell = currentRow.createCell(cellIndex)
            when (value) {
                is String -> {
                    cell.setCellValue(value)
                }
                is Double -> {
                    cell.setCellValue(value)
                }
                is LocalDateTime -> {
                    cell.setCellValue(value)
                }
                is BigDecimal -> {
                    cell.setCellValue(value.toDouble())
                }
            }
        }
    }
    return workBook
}

fun main() {
    val workBook = setUpWorkBook()

    //sort

    //pivot

    //aggregate

    //filter


    FileOutputStream("test.xlsx").use {
        workBook.write(it)
    }
}
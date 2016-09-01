package co.uk.taurasystems.utils

import co.uk.taurasystems.utils.ExcelDocHelper.Companion.getCellValueAsString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.util.*

/**
 * Created by tauraamui on 01/09/2016.
 */
class ExcelDocHelper {

    companion object {

        //Extension function for Cell
        fun Cell.getCellValueAsString(): String {
            when (this.cellType) {
                Cell.CELL_TYPE_STRING -> {
                    return this.richStringCellValue.toString()
                }
                Cell.CELL_TYPE_BOOLEAN -> {
                    return this.booleanCellValue.toString()
                }
                Cell.CELL_TYPE_ERROR -> {
                    return this.errorCellValue.toString()
                }
                Cell.CELL_TYPE_NUMERIC -> {
                    return this.numericCellValue.toString()
                }
                Cell.CELL_TYPE_BLANK -> {
                    return ""
                }
                Cell.CELL_TYPE_FORMULA -> {
                    return getFormulaResultValueAsString(this)
                }
            }
            return ""
        }

        fun Cell.getFormulaResultValueAsString(cell: Cell): String {
            when (cell.cachedFormulaResultType) {
                Cell.CELL_TYPE_STRING -> {
                    return cell.richStringCellValue.toString()
                }
                Cell.CELL_TYPE_BOOLEAN -> {
                    return cell.booleanCellValue.toString()
                }
                Cell.CELL_TYPE_ERROR -> {
                    return cell.errorCellValue.toString()
                }
                Cell.CELL_TYPE_NUMERIC -> {
                    return cell.numericCellValue.toString()
                }
                Cell.CELL_TYPE_BLANK -> {
                    return ""
                }
                Cell.CELL_TYPE_FORMULA -> {
                    return getFormulaResultValueAsString(cell)
                }
            }
            return ""
        }

        fun outputPopulatedCellData(file: File, sheetIndex: Int) {
            if (file.exists()) {
                val workbook = XSSFWorkbook(FileInputStream(file))
                val sheet = workbook.getSheetAt(sheetIndex)

                for (row in sheet) {
                    for (cell in row) {
                        print("\t\t${cell.getCellValueAsString()}")
                    }
                    println()
                }
                workbook.close()
            }
        }

        fun getDataFromExcelSheetColumn(file: File, sheetIndex: Int, columnIndex: Int): ArrayList<String> {
            val columnData = ArrayList<String>()
            if (file.exists()) {
                val workbook = XSSFWorkbook(FileInputStream(file))
                val sheet = workbook.getSheetAt(sheetIndex)
                for (row in sheet) {
                    columnData.add(row.getCell(columnIndex).getCellValueAsString())
                }
                workbook.close()
            }
            return columnData
        }

        fun getDataFromExcelSheetColumn(file: File, sheetIndex: Int, columnNamesRowIndex: Int, columnName: String): ArrayList<String> {
            var columnIndex = -1
            val data = ArrayList<String>()

            if (file.exists()) {
                val workbook = XSSFWorkbook(FileInputStream(file))
                val sheet = workbook.getSheetAt(sheetIndex)
                val columnNamesRow = sheet.getRow(columnNamesRowIndex)
                for (cell in columnNamesRow) {
                    if (cell.getCellValueAsString() == columnName) {
                        columnIndex = cell.columnIndex
                    }
                }
                workbook.close()
            }
            if (columnIndex >= 0) {
                return getDataFromExcelSheetColumn(file, sheetIndex, columnIndex)
            }
            return data
        }

        fun getExcelSheetIndex(file: File, sheetName: String): Int {
            if (file.exists()) {
                val workbook = XSSFWorkbook(FileInputStream(file))
                val sheetIndex = workbook.getSheetIndex(sheetName)
                workbook.close()
                return sheetIndex
            }
            return -1
        }
    }
}
package co.uk.taurasystems.utils

import co.uk.taurasystems.utils.ExcelDocHelper.Companion.getCellValueAsString
import org.apache.poi.hssf.usermodel.HSSFWorkbook
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

        private var modernWorkbook = XSSFWorkbook()
        private var modernWorkbookOpen = false

        private var legacyWorkbook = HSSFWorkbook()
        private var legacyWorkbookOpen = false

        enum class WorkbookType {MODERN, LEGACY }

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

        fun openWorkbook(file: File) {
            closeWorkbook()
            if (file.exists()) {
                if (FileHelper.getFileExt(file) == "xlsx") {
                    openModernWorkbook(file)
                } else if (FileHelper.getFileExt(file) == "xls") {
                    openLegacyWorkbook(file)
                } else {
                    //TODO: Need to change exception type to something more relevant
                    throw Exception("Incorrect file extension for Excel workbooks")
                }
            }
        }

        fun closeWorkbook() {
            closeModernWorkbook()
            closeLegacyWorkbook()
        }

        private fun openModernWorkbook(file: File) {
            modernWorkbook = XSSFWorkbook(FileInputStream(file))
            modernWorkbookOpen = true
        }

        private fun closeModernWorkbook() {
            if (modernWorkbookOpen) {
                modernWorkbook.close()
                modernWorkbookOpen = false
            }
        }

        private fun openLegacyWorkbook(file: File) {
            legacyWorkbook = HSSFWorkbook(FileInputStream(file))
            legacyWorkbookOpen = true
        }

        private fun closeLegacyWorkbook() {
            if (legacyWorkbookOpen) {
                legacyWorkbook.close()
                legacyWorkbookOpen = false
            }
        }

        fun outputPopulatedCellDataInSheet(sheetIndex: Int) {
            if (legacyWorkbookOpen) {
                val sheet = legacyWorkbook.getSheetAt(sheetIndex)
                for (row in sheet) {
                    for (cell in row) {
                        print("\t${cell.getCellValueAsString()}")
                    }
                    println()
                }
            } else if (modernWorkbookOpen) {
                val sheet = modernWorkbook.getSheetAt(sheetIndex)
                for (row in sheet) {
                    for (cell in row) {
                        print("\t${cell.getCellValueAsString()}")
                    }
                    println()
                }
            }
        }

        fun getDataFromExcelSheetColumn(sheetIndex: Int, columnIndex: Int): ArrayList<String> {
            val columnData = ArrayList<String>()
            if (legacyWorkbookOpen) {
                val sheet = legacyWorkbook.getSheetAt(sheetIndex)
                for (row in sheet) {
                    columnData.add(row.getCell(columnIndex).getCellValueAsString())
                }
            } else if (modernWorkbookOpen) {
                val sheet = modernWorkbook.getSheetAt(sheetIndex)
                for (row in sheet) {
                    columnData.add(row.getCell(columnIndex).getCellValueAsString())
                }
            }
            return columnData
        }

        fun getDataFromExcelSheetColumn(sheetIndex: Int, columnNamesRowIndex: Int, columnName: String): ArrayList<String> {
            var columnIndex = -1
            val data = ArrayList<String>()

            if (legacyWorkbookOpen) {
                val sheet = legacyWorkbook.getSheetAt(sheetIndex)
                val columnNamesRow = sheet.getRow(columnNamesRowIndex)
                for (cell in columnNamesRow) {
                    if (cell.getCellValueAsString() == columnName) {
                        columnIndex = cell.columnIndex
                    }
                }
            } else if (modernWorkbookOpen) {
                val sheet = modernWorkbook.getSheetAt(sheetIndex)
                val columnNamesRow = sheet.getRow(columnNamesRowIndex)
                for (cell in columnNamesRow) {
                    if (cell.getCellValueAsString() == columnName) {
                        columnIndex = cell.columnIndex
                    }
                }
            }
            if (columnIndex >= 0) {
                return getDataFromExcelSheetColumn(sheetIndex, columnIndex)
            }
            return data
        }

        fun getExcelSheetIndex(sheetName: String): Int {
            if (legacyWorkbookOpen) {
                return legacyWorkbook.getSheetIndex(sheetName)
            } else if (modernWorkbookOpen) {
                return modernWorkbook.getSheetIndex(sheetName)
            }
            return -1
        }
    }
}
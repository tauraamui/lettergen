package co.uk.taurasystems.utils;

import jdk.nashorn.internal.runtime.ScriptRuntime
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hwpf.HWPFDocument
import org.apache.poi.poifs.filesystem.NotOLE2FileException
import org.apache.poi.poifs.filesystem.OfficeXmlFileException
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xwpf.usermodel.TextSegement
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.IOException
import java.util.*

/**
 * Created by tauraaamui on 15/08/2016.
 */
class DocHelper {

    companion object {

        var modernWordDoc = XWPFDocument()
        var modernExcelWorkbook = XSSFWorkbook()

        fun openModernWorkDoc(file: File?) {
            if (isDocx(file)) {
                modernWordDoc = XWPFDocument(FileInputStream(file))
            }
        }

        fun openExcelWorkbook(file: File?) {
            if (isModernExcelDoc(file)) {

            }
        }

        fun documentTitleContains(file: File, textToFind: String): Boolean {
            if (isDoc(file) && file.name.contains(textToFind)) return true
            else if (isDocx(file) && file.name.contains(textToFind)) return true
            return false
        }

        fun documentTitleContains(file: File, textToFind: String, toLower: Boolean): Boolean {
            if (!toLower) return documentTitleContains(file, textToFind)
            if (isDoc(file) && file.name.toLowerCase().contains(textToFind.toLowerCase())) return true
            else if (isDocx(file) && file.name.toLowerCase().contains(textToFind.toLowerCase())) return true
            return false
        }

        fun documentContains(file: File, textToFind: String): Boolean {
            if (isDoc(file)) {
                try {
                    val hwpfDocument = HWPFDocument(FileInputStream(file))
                    for (i in 0..hwpfDocument.range.numParagraphs()-1) {
                        if (hwpfDocument.range.getParagraph(i).text().toLowerCase().contains(textToFind.toLowerCase())) return true
                    }
                } catch (e: Exception) {
                    e.printStackTrace()
                }
            } else if (isDocx(file)) {
                try {
                    //TODO: Check to make sure this is actually they way to do string searching within a docx file...
                    val xwpfDocument = XWPFDocument(FileInputStream(file))
                    for (paragraph in xwpfDocument.paragraphsIterator) {
                        if (paragraph.text.contains(textToFind)) return true
                    }
                } catch (e: Exception) {
                    e.printStackTrace()
                }
            }
            return false
        }

        fun isDoc(file: File?): Boolean {
            if (getFileExt(file) == "doc") return true
            return false
        }

        fun isDocx(file: File?): Boolean {
            if (getFileExt(file) == "docx") return true
            return false
        }

        fun isModernExcelDoc(file: File?): Boolean {
            if (getFileExt(file) == "xlsx") return true
            return false
        }

        fun isLegacyExcelDoc(file: File?): Boolean {
            if (getFileExt(file) == "xls") return true
            return false
        }

        fun findAndReplaceTagsInWordDoc(file: File?, keysAndValues: HashMap<String, String?>): Any? {
            if (isDoc(file)) {
                return replaceTagsInLegacyWordDoc(file, keysAndValues)
            } else if (isDocx(file)) {
                return replaceTagsInModernWordDoc(file, keysAndValues)
            } else {
                throw Exception("extension must match either '.doc' or '.docx'")
            }
        }

        fun replaceTagsInLegacyWordDoc(file: File?, keysAndValues: HashMap<String, String?>): HWPFDocument? {
            var hwpfDocument: HWPFDocument? = null
            try {
                hwpfDocument = HWPFDocument(FileInputStream(file))
                val range = hwpfDocument.range
                for (i in 0.. range.numParagraphs()-1) {
                    val paragraph = range.getParagraph(i)
                    for (j in 0..paragraph.numCharacterRuns()-1) {
                        val run = paragraph.getCharacterRun(j)
                        for ((key, value) in keysAndValues) {
                            run.replaceText(key, value)
                        }
                    }
                }
                hwpfDocument.dataStream.inputStream().close()
            } catch (e: OfficeXmlFileException) {
                println("Document ${file?.name} is a newer .docx format...")
            } catch (e: NotOLE2FileException) {
                println("Document ${file?.name} has an invalid header signature")
            } catch (e: IOException) {
                e.printStackTrace()
            }
            return hwpfDocument
        }

        fun replaceTagsInModernWordDoc(file: File?, keysAndValues: HashMap<String, String?>): XWPFDocument? {
            var xwpfDocument: XWPFDocument? = null
            try {
                xwpfDocument = XWPFDocument(FileInputStream(file))
                xwpfDocument.paragraphs.forEach {
                    println(it.text)
                }
                xwpfDocument.close()
            } catch (e: IOException) {
                e.printStackTrace()
            }
            return xwpfDocument
        }

        fun getFileExt(file: File?): String {
            if (file == null) return ""
            if (file.name.contains(".")) {
                return file.name.split(".")[1]
            } else {
                return ""
            }
        }

        fun outputPopulatedExcelSheetCellContents(file: File?, sheetIndex: Int) {
            if (file?.exists()!!) {
                val workbook = XSSFWorkbook(FileInputStream(file))
                val sheetToReadFrom = workbook.getSheetAt(sheetIndex)

                for (row in sheetToReadFrom) {
                    for (cell in row) {
                        if (cell.cellType == XSSFCell.CELL_TYPE_NUMERIC) {
                            print("\t\t${cell.numericCellValue}")
                        } else if (cell.cellType == XSSFCell.CELL_TYPE_FORMULA) {
                            if (cell.cachedFormulaResultType == XSSFCell.CELL_TYPE_NUMERIC) {
                                print("\t\t${cell.numericCellValue}")
                            }
                        }
                    }
                    println()
                }
            }
        }

        fun getDataFromExcelSheetColumn(file: File?, sheetIndex: Int, columnIndex: Int): ArrayList<String> {
            val columnData = ArrayList<String>()
            if (file?.exists()!!) {
                val workbook = XSSFWorkbook(FileInputStream(file))
                val sheet = workbook.getSheetAt(sheetIndex)
                for (row in sheet) {
                    columnData.add(row.getCell(columnIndex).numericCellValue.toString())
                }
            }
            return columnData
        }

        fun getDataFromExcelSheetColumn(file: File?, sheetIndex: Int, columnName: String) {
            var columnIndex = -1
            val workbook = XSSFWorkbook(FileInputStream(file))
        }
    }
}
package co.uk.taurasystems

import co.uk.taurasystems.utils.ExcelDocHelper
import co.uk.taurasystems.utils.WordDocHelper
import co.uk.taurasystems.utils.FileHelper
import java.io.File
import java.io.FileOutputStream

/**
 * Created by alewis on 01/09/2016.
 */

fun main(args: Array<String>) {
    val workbook = File("Book1.xlsx")
    //Dochelper.outputPopulatedExcelSheetCellContents(workbook, 0)

    val sheetIndex = ExcelDocHelper.getExcelSheetIndex(workbook, "Sheet1")
    if (sheetIndex < 0) return
    /*
    for (cellData in ExcelDocHelper.getDataFromExcelSheetColumn(workbook, sheetIndex, 0, "NT60/ps")) {
        println(cellData)
    }
    */

    ExcelDocHelper.outputPopulatedCellData(workbook, ExcelDocHelper.getExcelSheetIndex(workbook, "Sheet1"))
}
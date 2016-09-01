package co.uk.taurasystems

import co.uk.taurasystems.utils.DocHelper
import co.uk.taurasystems.utils.FileHelper
import java.io.File
import java.io.FileOutputStream

/**
 * Created by alewis on 01/09/2016.
 */

fun main(args: Array<String>) {
    val workbook = File("Book1.xlsx")
    val newFileToSave = File(FileHelper.getUniqueFileName(workbook))
    val fos = FileOutputStream(newFileToSave)
    fos.close()

    //Dochelper.outputPopulatedExcelSheetCellContents(workbook, 0)

    for (cellData in DocHelper.getDataFromExcelSheetColumn(workbook, 0, 2)) {
        println(cellData)
    }
}
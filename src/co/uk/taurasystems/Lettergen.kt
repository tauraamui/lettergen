package co.uk.taurasystems

import co.uk.taurasystems.utils.ExcelDocHelper
import co.uk.taurasystems.utils.WordDocHelper
import org.apache.poi.ss.formula.functions.T
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileOutputStream
import java.util.*

/**
 * Created by alewis on 01/09/2016.
 */

class Lettergen {

    var excelFilePath = ""
    var letterTemplateFilePath = ""
    val recipientsList = arrayListOf<Recipient>()

    fun init() {
        setupRecipients()
    }

    fun setupRecipients() {
        val excelDocument = File(excelFilePath)
        if (excelDocument.exists()) {
            ExcelDocHelper.openWorkbook(excelDocument)
            mapExcelDataToRecipients()
            outputLetterTemplateContent()
        }
    }

    fun mapExcelDataToRecipients() {
        val memberIDs = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "mem")
        val recipientTitles = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Title")
        val recipientFirstnames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Given")
        val familyNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Family")
        val addressees = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Addressee")
        val otherPerson = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Other Person")
        val houseNumberNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "s#")
        val streetNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Street")
        val townNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Town")
        val postCodes = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Postcode")
        val telephoneNumbers = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Tel")
        val emailAddresses = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Email")

        for (i in 0..memberIDs.size-1) {
            recipientsList.add(Recipient(fromStringToFloatToLong(memberIDs[i]), recipientTitles[i].trim(), recipientFirstnames[i].trim(),
                                         familyNames[i].trim(), addressees[i].trim(), fromStringToFloatToLong(otherPerson[i].trim()),
                                            formatHouseNumberString(houseNumberNames[i]).replace(".", ""), streetNames[i].trim(), townNames[i].trim(), postCodes[i].trim(), telephoneNumbers[i].trim(), emailAddresses[i].trim()))
        }
        recipientsList.removeIf { it.memberID < 0 }

        recipientsList.forEach { println(it) }
    }

    fun outputLetterTemplateContent() {
        val map = HashMap<String, String?>()
        map.putIfAbsent("<Address1Line1>", "Testing")
        WordDocHelper.replaceTagsInModernWordDoc(File(letterTemplateFilePath), map)
    }

    fun formatHouseNumberString(stringToFormat: String): String {
        if (stringToFormat.contains(".")) {
            return stringToFormat.substring(0, stringToFormat.lastIndexOf("."))
        }
        return stringToFormat
    }

    fun fromStringToFloatToLong(string: String): Long {
        try {
            val converted = string.toFloat().toLong()
            return converted
        } catch (e: Exception) {
            return -1
        }
    }
}

fun main(args: Array<String>) {
    if (args.size >= 1) {
        val letterGen = Lettergen()
        letterGen.excelFilePath = args[0]
        letterGen.letterTemplateFilePath = args[1]
        letterGen.init()
    }
}
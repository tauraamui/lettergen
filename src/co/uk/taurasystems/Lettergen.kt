package co.uk.taurasystems

import co.uk.taurasystems.utils.ExcelDocHelper
import org.apache.poi.ss.formula.functions.T
import java.io.File
import java.util.*

/**
 * Created by alewis on 01/09/2016.
 */

fun main(args: Array<String>) {
    if (args.isNotEmpty()) {

        val workbook = File(args[0])
        ExcelDocHelper.openWorkbook(workbook)

        val memberIDs = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "mem")
        val recipientTitles = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Title")
        val recipientFirstname = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Given")
        val familyNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Family")
        val otherPerson = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Other Person")
        val addressees = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Addressee")

        val recipientList = ArrayList<Recipient>()

        val listsSizes = memberIDs.size
        if (recipientTitles.size == listsSizes && recipientFirstname.size == listsSizes && familyNames.size == listsSizes && otherPerson.size == listsSizes && addressees.size == listsSizes) {
            for (i in 0..listsSizes-1) {
                val recipient = Recipient(-1, "", "", "", "", "")
                var idToSet = -1
                try {
                    val memberIDAsFloat = memberIDs[i].toFloat()
                    idToSet = memberIDAsFloat.toInt()
                } catch (e: Exception) {}
                recipient.memberID = idToSet
                recipient.title = recipientTitles[i]
                recipient.firstName = recipientFirstname[i]
                recipient.surname = familyNames[i]
                recipient.addressee = addressees[i]
                recipient.otherPerson = otherPerson[i]
                recipientList.add(recipient)
            }
            recipientList.removeIf {
                it.memberID < 0
            }
        }

        for (recipient in recipientList) {
            println(recipient)
        }

        ExcelDocHelper.closeWorkbook()
    }
}
package co.uk.taurasystems

import co.uk.taurasystems.utils.ExcelDocHelper
import org.apache.poi.ss.formula.functions.T
import java.io.File
import java.util.*

/**
 * Created by alewis on 01/09/2016.
 */

var memberIDs = null
var recipientTitles = null
var recipientFirstname = null
var familyNames = null
var otherPerson = null
var addressees = null

class Lettergen {

    private var args = arrayOf<String>()
    private var memberIDs: Array<String>? = null
    private var recipientTitles: Array<String>? = null
    private var recipientFirstname: Array<String>? = null
    private var familyNames: Array<String>? = null
    private var otherPerson: Array<String>? = null
    private var addressees: Array<String>? = null
    private var houseNameNumbers: Array<String> = arrayOf<String>()
    private val recipientList: Array<Recipient> = arrayOf<Recipient>()

    constructor(args: Array<String>) {
        this.args = args
    }

    fun start() {
        if (args.isNotEmpty()) {
            val workbook = File(args[0])
            ExcelDocHelper.openWorkbook(workbook)
            memberIDs = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "mem")
            recipientTitles = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Title")
            recipientFirstname = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Given")
            familyNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Family")
            otherPerson = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Other Person")
            addressees = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Addressee")
            houseNameNumbers = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "s#")

            val recipientList = ArrayList<Recipient>()
            val listsSizes = memberIDs?.size
            if (recipientTitles?.size == listsSizes && recipientFirstname?.size == listsSizes && familyNames?.size == listsSizes && otherPerson?.size == listsSizes && addressees?.size == listsSizes) {
                for (i in 0..listsSizes!!-1) {
                    val recipient = Recipient(-1, "", "", "", -1, "")
                    var idToSet: Long = -1
                    idToSet = stringToFloatToLong(memberIDs!![i])
                    recipient.memberID = idToSet
                    recipient.title = recipientTitles!![i]
                    recipient.firstName = recipientFirstname!![i]
                    recipient.surname = familyNames!![i]
                    recipient.otherPerson = stringToFloatToLong(otherPerson!![i])
                    recipient.addressee = addressees!![i]
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
}

fun main(args: Array<String>) {
    val letterGen = Lettergen(args)
    letterGen.start()
    /*
    if (args.isNotEmpty()) {
        val workbook = File(args[0])
        ExcelDocHelper.openWorkbook(workbook)
        val memberIDs = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "mem")
        val recipientTitles = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Title")
        val recipientFirstname = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Given")
        val familyNames = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Family")
        val otherPerson = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Other Person")
        val addressees = ExcelDocHelper.getDataFromExcelSheetColumn(0, 0, "Addressee")
        val

        val recipientList = ArrayList<Recipient>()

        val listsSizes = memberIDs.size
        if (recipientTitles.size == listsSizes && recipientFirstname.size == listsSizes && familyNames.size == listsSizes && otherPerson.size == listsSizes && addressees.size == listsSizes) {
            for (i in 0..listsSizes-1) {
                val recipient = Recipient(-1, "", "", "", -1, "")
                var idToSet: Long = -1
                idToSet = stringToFloatToLong(memberIDs[i])
                recipient.memberID = idToSet
                recipient.title = recipientTitles[i]
                recipient.firstName = recipientFirstname[i]
                recipient.surname = familyNames[i]
                recipient.addressee = addressees[i]
                recipient.otherPerson = stringToFloatToLong(otherPerson[i])
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
    */
}

private fun listsAreSameLength(listLength: Int) {

}

fun stringToFloatToLong(number: String): Long {
    var otherId: Long = -1
    try {
        val otherPersonID = number.toFloat()
        otherId = otherPersonID.toLong()
    } catch (e: Exception) {}
    return otherId
}
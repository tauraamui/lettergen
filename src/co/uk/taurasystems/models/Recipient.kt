package co.uk.taurasystems.models

/**
 * Created by alewis on 02/09/2016.
 */
data class Recipient(var memberID: Long, var title: String, var firstName: String,
                     var surname: String, var addressee: String, var otherPerson: Long,
                     var houseNumberName: String, var streetName: String, var townName: String,
                     var postCode: String, var telephoneNumber: String, var emailAddress: String)
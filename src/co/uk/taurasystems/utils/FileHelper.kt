package co.uk.taurasystems.utils

import java.io.File

/**
 * Created by alewis on 01/09/2016.
 */
class FileHelper {

    companion object {

        fun getUniqueFileName(file: File): String {
            var extension = ""

            if (file.absolutePath.contains(".")) {
                extension = file.absolutePath.substring(file.absolutePath.lastIndexOf("."), file.absolutePath.length).replace(".", "")
                val fileWithoutExt = File(file.absolutePath.substring(0, file.absolutePath.lastIndexOf(".")).replace(".", ""))
                return getUniqueFileNameSepExt(fileWithoutExt, extension)
            } else {
                var fileToSave = file
                var versionSuffix = 1

                if (!file.exists()) return file.absolutePath
                fileToSave = File("${file.absolutePath} $versionSuffix")

                while (fileToSave.exists()) {
                    fileToSave = File("${file.absolutePath} $versionSuffix")
                    versionSuffix++
                }
                return fileToSave.absolutePath
            }
        }

        fun getUniqueFileNameSepExt(file: File, extension: String): String {

            var fileToSave = file
            var versionSuffix = 1

            val firstFile = File(file.absolutePath + "." + extension)
            if (!firstFile.exists()) return firstFile.absolutePath

            fileToSave = File(file.absolutePath + " $versionSuffix." + extension)

            while (fileToSave.exists()) {
                fileToSave = File(file.absolutePath + " $versionSuffix." + extension)
                versionSuffix++
            }
            return fileToSave.absolutePath
        }
    }
}

package org.example

fun main() {
    val dataProcessor = DataProcessor()
    print("File path: ")
    val filePath = "strings_with_translate.xlsx"//readln()
    val jsonsResult = dataProcessor.parseXlsxToJson(filePath)
//    print("Strings from xlsx as json:\n${jsonsResult?.strings}\n")
//    print("Arrays from xlsx as json:\n${jsonsResult?.arrays}\n")
    dataProcessor.stringsToXmlFile(jsonsResult?.strings, "E:\\WB\\strings_example.xml")
    dataProcessor.arraysToXmlFile(jsonsResult?.arrays, "E:\\WB\\arrays_example.xml")
}
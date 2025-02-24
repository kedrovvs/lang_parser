package org.example

import org.apache.poi.ss.usermodel.*
import org.json.JSONObject
import org.json.JSONTokener
import java.io.File
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.transform.OutputKeys
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult

class DataProcessor {

    fun parseXlsxToJson(filePath: String): JsonsFromXlsx? {
        try {
            val file = File(filePath)
            if (!file.exists()) {
                print("File not found: $filePath\n")
                return null
            }

            val workbook = WorkbookFactory.create(file)
            val sheet = workbook.getSheetAt(0)

            val headers = mutableListOf<String>()
            val headerRow = sheet.getRow(0)
            for (i in 0 until headerRow.physicalNumberOfCells) {
                val cell = headerRow.getCell(i) ?: continue
                headers.add(cell.stringCellValue.trim())
            }
            for (header in headers.subList(2, headers.size)) {
                print("$header ${headers.indexOf(header)}\n")
            }
            print("Choose language: ")
            val chosenLanguageIndex: Int?
            try {
                chosenLanguageIndex = readln().toInt()
                if (chosenLanguageIndex > headers.size || chosenLanguageIndex < 2) {
                    print("Out of range.\n")
                    return null
                }
            } catch (e: NumberFormatException) {
                print(e.message.toString())
                return null
            }

            val stringsAsJsonObject = JSONObject()
            val arraysAsJsonObject = JSONObject()
            var whereArraysStartedRowIndex = -1

            //Strings processing
            for (rowIndex in 1..sheet.lastRowNum) {
                val row = sheet.getRow(rowIndex)
                if (row == null) {
                    whereArraysStartedRowIndex = rowIndex + 2
                    break
                }

                val identifier = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.trim()
                val cell = row.getCell(
                    chosenLanguageIndex,
                    Row.MissingCellPolicy.CREATE_NULL_AS_BLANK
                )

                val cellValue = when (cell.cellType) {
                    CellType.STRING -> cell.stringCellValue.trim()
                    CellType.NUMERIC -> {
                        if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                            cell.dateCellValue.toString() // Форматирование даты по необходимости
                        } else {
                            cell.numericCellValue.toString() // Преобразуем число в строку
                        }
                    }

                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                    CellType.BLANK -> "" // Обработка пустых ячеек
                    CellType.FORMULA -> {
                        try {
                            cell.stringCellValue.trim()
                        } catch (e: IllegalStateException) {
                            cell.numericCellValue.toString()
                        }
                    }

                    else -> "" // Обработка других типов ячеек
                }
                stringsAsJsonObject.put(identifier, cellValue)
            }

            //Arrays processing
            if (whereArraysStartedRowIndex > 0) {
                var currentArrayName: String?
                val arrayValues = mutableListOf<String>()

                for (rowIndex in whereArraysStartedRowIndex..sheet.lastRowNum) {
                    val row = sheet.getRow(rowIndex) ?: break

                    val identifierCell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                    currentArrayName = getCellValueAsString(identifierCell).takeIf { it.isNotEmpty() }
                    currentArrayName ?: continue
                    val mergedRegionSize = getMergedRegionSize(sheet, identifierCell) ?: continue

                    for (cellIndex in mergedRegionSize.first..mergedRegionSize.second) {
                        val valueRow = sheet.getRow(cellIndex)
                        val stringValueCell = getCellValueAsString(
                            valueRow.getCell(
                                chosenLanguageIndex,
                                Row.MissingCellPolicy.CREATE_NULL_AS_BLANK
                            )
                        )
                        arrayValues.add(stringValueCell)
                    }

                    arraysAsJsonObject.put(currentArrayName, arrayValues)
                    arrayValues.clear()
                }
            }

            workbook.close()
            return JsonsFromXlsx(
                strings = stringsAsJsonObject.toString(2),
                arrays = arraysAsJsonObject.toString(2)
            )

        } catch (e: Exception) {
            print("Error parsing xlsx file: ${e.message}\n")
            e.printStackTrace()
            return null
        }
    }

    fun stringsToXmlFile(stringsAsJson: String?, filePath: String) {
        if (stringsAsJson == null) {
            print("Source json is empty, stopping execution.\n")
            return
        }
        try {
            val jsonObject = JSONObject(JSONTokener(stringsAsJson))

            val docFactory = DocumentBuilderFactory.newInstance()
            val docBuilder = docFactory.newDocumentBuilder()

            val newStringsXmlDoc = docBuilder.newDocument()
            val rootElement = newStringsXmlDoc.createElement("resources")
            rootElement.setAttribute("xmlns:tools","http://schemas.android.com/tools")
            newStringsXmlDoc.appendChild(rootElement)

            jsonObject.keys().forEach { stringIdentifier ->
                val stringValue = jsonObject.getString(stringIdentifier)

                val stringAsXmlItem = newStringsXmlDoc.createElement("string")
                stringAsXmlItem.setAttribute("name", stringIdentifier)
                stringAsXmlItem.appendChild(newStringsXmlDoc.createTextNode(stringValue))
                rootElement.appendChild(stringAsXmlItem)
            }

            val transformerFactory = TransformerFactory.newInstance()
            val transformer = transformerFactory.newTransformer()

            transformer.setOutputProperty(OutputKeys.INDENT, "yes")
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2")

            val source = DOMSource(newStringsXmlDoc)
            val result = StreamResult(File(filePath))
            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes")
            transformer.transform(source, result)

        } catch (e: Exception) {
            print("Error trying transfer strings to xml-file:\n${e.message.toString()}\n")
        }
    }

    fun arraysToXmlFile(arraysAsJson: String?, filePath: String) {
        if (arraysAsJson == null) {
            print("Source json is empty, stopping execution.\n")
            return
        }
        try {
            val jsonObject = JSONObject(JSONTokener(arraysAsJson))

            val docFactory = DocumentBuilderFactory.newInstance()
            val docBuilder = docFactory.newDocumentBuilder()
            val newStringsXmlDoc = docBuilder.newDocument()

            val rootElement = newStringsXmlDoc.createElement("resources")
            rootElement.setAttribute("xmlns:tools","http://schemas.android.com/tools")
            newStringsXmlDoc.appendChild(rootElement)

            jsonObject.keys().forEach { stringIdentifier ->
                val array = jsonObject.getJSONArray(stringIdentifier).toList()
                print("$array\n")

                val stringArray = newStringsXmlDoc.createElement("string-array")
                stringArray.setAttribute("name", stringIdentifier)
                for (stringValue in array) {
                    val newItem = newStringsXmlDoc.createElement("item")
                    newItem.appendChild(newStringsXmlDoc.createTextNode(stringValue.toString()))
                    stringArray.appendChild(newItem)
                }
                rootElement.appendChild(stringArray)
            }

            val transformerFactory = TransformerFactory.newInstance()
            val transformer = transformerFactory.newTransformer()

            transformer.setOutputProperty(OutputKeys.INDENT, "yes")
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2")

            val source = DOMSource(newStringsXmlDoc)
            val result = StreamResult(File(filePath))

            transformer.setOutputProperty(OutputKeys.STANDALONE, "yes")
            transformer.transform(source, result)

        } catch (e: Exception) {
            print("Error trying transfer arrays to xml-file:\n${e.message.toString()}\n")
        }
    }


    private fun getCellValueAsString(cell: Cell?): String {
        return when (cell?.cellType) {
            CellType.STRING -> cell.stringCellValue.trim()
            CellType.NUMERIC -> cell.numericCellValue.toString().trim()
            CellType.BOOLEAN -> cell.booleanCellValue.toString().trim()
            CellType.BLANK -> ""
            else -> ""
        }
    }

    private fun getMergedRegionSize(sheet: Sheet, cell: Cell?): Pair<Int, Int>? {
        if (cell == null) return null

        for (i in 0 until sheet.numMergedRegions) {
            val mergedRegion = sheet.getMergedRegion(i)
            if (mergedRegion.isInRange(cell.rowIndex, cell.columnIndex)) {
                return Pair(mergedRegion.firstRow, mergedRegion.lastRow)
            }
        }
        return null
    }
}
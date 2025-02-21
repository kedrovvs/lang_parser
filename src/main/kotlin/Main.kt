package org.example

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.json.JSONObject
import java.io.File

fun main() {
    print("File path: ")
    val filePath = "E:\\WB\\strings_with_translate.xlsx"//readln()
    val jsonString = parseXlsxToJson(filePath)
    print("json_string: $jsonString")
}

fun parseXlsxToJson(filePath: String): String? {
    try {
        val file = File(filePath)
        if (!file.exists()) {
            print("File not found: $filePath\n")
            return null // Возвращаем пустую строку, если файл не найден.  Можно выбросить исключение, если это предпочтительнее.
        }

        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.getSheetAt(0) // Берем первый лист. Можно изменить, если нужно.

        val headers = mutableListOf<String>() // Список для хранения заголовков столбцов.

        // Получаем заголовки из первой строки
        val headerRow = sheet.getRow(0)
        for (i in 0 until headerRow.physicalNumberOfCells) { //  Используем `physicalNumberOfCells` для правильной обработки пустых ячеек.
            val cell = headerRow.getCell(i) ?: continue // Обработка пустой ячейки в заголовке.
            headers.add(cell.stringCellValue.trim()) // Добавляем заголовок, убираем лишние пробелы
        }
        for (header in headers.subList(2, headers.size)) {
            print("$header ${headers.indexOf(header)}\n")
        }
        print("Choose language: ")
        val chosenLanguageIndex: Int?
        val chosenLanguage: String?
        try {
            chosenLanguageIndex = readln().toInt()
            if (chosenLanguageIndex > headers.size || chosenLanguageIndex < 2) {
                print("Out of range.\n")
                return null
            }
            chosenLanguage = headers[chosenLanguageIndex]
        } catch (e: NumberFormatException) {
            print(e.message.toString())
            return null
        }

        // Обрабатываем строки данных, начиная со второй строки (пропускаем заголовки)
        val jsonObject = JSONObject() // Объект JSON для каждой строки

        for (rowIndex in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue // Пропускаем пустые строки

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
            jsonObject.put(identifier, cellValue) // Добавляем значение в JSON объект
        }
        val result = JSONObject().put(chosenLanguage, jsonObject)

        workbook.close() // Закрываем workbook после использования
        return result.toString(2) // Форматируем JSON для удобочитаемости (отступы = 2)

    } catch (e: Exception) {
        print("Error parsing xlsx file: ${e.message}\n")
        e.printStackTrace()
        return null
    }
}

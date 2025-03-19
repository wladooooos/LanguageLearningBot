import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row

object Musor {
    fun checkUserState(chatId: Long, filePath: String, sheetName: String = "Состояние пользователя"): Boolean {
        println("WWW checkUserState // Проверка состояния пользователя")
        println("W1 Вход в функцию. Параметры: chatId=$chatId, filePath=\"$filePath\", sheetName=\"$sheetName\"")

        val excelManager = ExcelManager(filePath)
        var allCompleted = false

        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                println("W3 ❌ Лист $sheetName не найден.")
                return@useWorkbook
            }
            println("W4 ✅ Лист найден. Проверка строк...")

            var userRow: Row? = null
            for (rowIndex in 1..sheet.lastRowNum.coerceAtLeast(1)) {
                val row = sheet.getRow(rowIndex) ?: sheet.createRow(rowIndex)
                val idCell = row.getCell(0) ?: row.createCell(0)

                val currentId = when (idCell.cellType) {
                    CellType.NUMERIC -> idCell.numericCellValue.toLong()
                    CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                    else -> null
                }
                println("W5 Проверяем строку ${rowIndex + 1}. ID: $currentId")

                if (currentId == chatId) {
                    userRow = row
                    println("W6 ✅ Пользователь найден в строке ${rowIndex + 1}")
                    break
                }

                if (currentId == null || currentId == 0L) {
                    println("W7 ⚠️ Пустая строка. Добавляем нового пользователя.")
                    idCell.setCellValue(chatId.toDouble())
                    for (i in 1..6) {
                        row.createCell(i).setCellValue(0.0)
                    }
                    excelManager.safelySaveWorkbook(workbook)
                    println("C8 ✅ Новый пользователь добавлен в строку ${rowIndex + 1}")
                    return@useWorkbook
                }
            }

            allCompleted = userRow?.let {
                (1..6).all { index ->
                    val cell = it.getCell(index)
                    val value = cell?.toString()?.toDoubleOrNull() ?: 0.0
                    println("W9 Проверка колонки $index. Значение: $value")
                    value > 0
                }
            } ?: false
        }

        println("W10 Все этапы завершены: $allCompleted")
        return allCompleted
    }
}
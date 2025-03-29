import com.github.kotlintelegrambot.config.Config
import org.apache.poi.ss.usermodel.CellType

fun updateUserBlockCompleted(chatId: Long, filePath: String) {
    println("Updating block completion status for chatId=$chatId")
    val excelManager = ExcelManager(filePath)
    excelManager.useWorkbook { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
            ?: throw IllegalArgumentException("Лист 'Состояние пользователя' не найден")
        var userRow: org.apache.poi.ss.usermodel.Row? = null
        // Явно получаем итератор по строкам листа
        for (row in sheet.iterator()) {
            // row теперь имеет тип org.apache.poi.ss.usermodel.Row
            val cell = row.getCell(0)
            val idValue = when (cell?.cellType) {
                CellType.NUMERIC -> cell.numericCellValue.toLong()
                CellType.STRING -> cell.stringCellValue.toLongOrNull()
                else -> null
            }
            if (idValue == chatId) {
                userRow = row
                break
            }
        }
        if (userRow == null) {
            println("Пользователь с chatId=$chatId не найден в листе 'Состояние пользователя'")
            return@useWorkbook
        }
        val blockCompletion = mutableListOf<Boolean>()
        // Обновляем для каждого блока (1, 2 и 3) согласно диапазону из Config.USER_STATE_BLOCK_RANGES
        for ((block, range) in Config.USER_STATE_BLOCK_RANGES) {
            var sum = 0.0
            for (col in range) {
                val cell = userRow.getCell(col)
                val value = when (cell?.cellType) {
                    CellType.NUMERIC -> cell.numericCellValue
                    CellType.STRING -> cell.stringCellValue.toDoubleOrNull() ?: 0.0
                    else -> 0.0
                }
                sum += value
            }
            blockCompletion.add(sum > 0)
        }
        if (blockCompletion.size == 3) {
            Globals.userBlockCompleted[chatId] =
                Triple(blockCompletion[0], blockCompletion[1], blockCompletion[2])
            println("Updated userBlockCompleted for chatId=$chatId: ${Globals.userBlockCompleted[chatId]}")
        }
    }
}

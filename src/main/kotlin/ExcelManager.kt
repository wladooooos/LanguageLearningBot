import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream

class ExcelManager(private val filePath: String) {

    private val file = File(filePath)

    // Функция-обёртка для работы с Workbook. Автоматически создаёт, передаёт в block и закрывает Workbook.
    fun <T> useWorkbook(block: (Workbook) -> T): T {
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        WorkbookFactory.create(file).use { workbook ->
            return block(workbook)
        }
    }

    // Функция для безопасного сохранения Workbook в файл
    fun safelySaveWorkbook(workbook: Workbook) {
        val tempFile = File("$filePath.tmp")
        try {
            FileOutputStream(tempFile).use { outputStream ->
                workbook.write(outputStream)
            }
            val originalFile = File(filePath)
            if (originalFile.exists() && !originalFile.delete()) {
                throw IllegalStateException("Не удалось удалить оригинальный файл: $filePath")
            }
            if (!tempFile.renameTo(originalFile)) {
                throw IllegalStateException("Не удалось переименовать временный файл: ${tempFile.path}")
            }
        } catch (e: Exception) {
            tempFile.delete()
            throw e
        }
    }

    fun ensureUserRecord(chatId: Long, filePath: String) {
        println("\nensureUserRecord запущен для chatId = $chatId, filePath = $filePath\n")
        val file = File(filePath)
        if (!file.exists()) {
            println("Ошибка: Файл $filePath не найден!")
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheetName = "Состояние пользователя"
            // Получаем лист или создаём его, если отсутствует
            val sheet = workbook.getSheet(sheetName) ?: run {
                println("Лист '$sheetName' не найден, создаём новый лист.")
                workbook.createSheet(sheetName)
            }
            println("Работаем с листом '$sheetName'. Начинаем поиск в столбце A, начиная со второй строки (индекс 1).")

            var recordFound = false
            var targetRowIndex: Int? = null
            // Итерируем строки, начиная со второй (индекс 1)
            for (i in 1..sheet.lastRowNum) {
                val row = sheet.getRow(i) ?: continue
                val cell = row.getCell(0)
                if (cell == null || cell.toString().trim().isEmpty()) {
                    println("Строка $i: ячейка пустая. Будет использована для новой записи.")
                    targetRowIndex = i
                    break
                } else {
                    val idValue = when (cell.cellType) {
                        CellType.NUMERIC -> cell.numericCellValue.toLong()
                        CellType.STRING -> cell.stringCellValue.trim().toLongOrNull()
                        else -> null
                    }
                    println("Строка $i: найден ID = $idValue")
                    if (idValue == chatId) {
                        println("Строка $i: ID совпадает с chatId = $chatId")
                        recordFound = true
                        targetRowIndex = i
                        break
                    }
                }
            }
            // Если ни одна подходящая строка не найдена, создаем новую запись в следующей строке после последней реально созданной строки
            if (targetRowIndex == null) {
                val lastRow = sheet.iterator().asSequence().maxByOrNull { it.rowNum }
                targetRowIndex = if (lastRow != null) lastRow.rowNum + 1 else 1
                println("Не найдено пустой строки, создаём новую запись в строке $targetRowIndex")
            }
            if (!recordFound) {
                // Создаем новую запись: записываем ID в столбец A и устанавливаем начальные значения 0 в остальных (6 падежей)
                val row = sheet.getRow(targetRowIndex) ?: sheet.createRow(targetRowIndex)
                row.createCell(0).setCellValue(chatId.toDouble())
                println("Записан chatId = $chatId в строке $targetRowIndex, колонка A")
                for (col in 1..6) {
                    row.createCell(col).setCellValue(0.0)
                    println("Установлено начальное значение 0.0 в колонке $col строки $targetRowIndex")
                }
                excelManager.safelySaveWorkbook(workbook)
                println("Создана новая запись для chatId = $chatId в строке $targetRowIndex")
            } else {
                println("Запись для chatId = $chatId уже существует в строке $targetRowIndex. Новая запись не создается.")
            }
        }
    }

}

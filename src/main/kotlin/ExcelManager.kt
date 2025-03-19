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
}

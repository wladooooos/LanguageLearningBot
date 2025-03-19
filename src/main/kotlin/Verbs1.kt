import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.keyboards.Keyboards
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import kotlin.collections.joinToString

object Verbs1 {
    fun callVerbs1(chatId: Long, bot: Bot) {
        println("callVerbs1: Обработка блока глаголов 1")
        Globals.userRandomVerb.remove(chatId)
        Globals.userBlocks[chatId] = 7
        Globals.userStates[chatId] = 0
        handleBlockVerbs1(chatId, bot)
    }

    fun handleBlockVerbs1(chatId: Long, bot: Bot) {
        println("aaa1 handleBlockVerbs1 // Вход в функцию. chatId=$chatId")
        val filePath = "Глаголы.xlsx"
        val sheetName = "Глаголы 1"
        println("aaa2 Параметры: filePath=$filePath, sheetName=$sheetName")

        // Всего нужно пройти колонки A..R (18 колонок: 0..17)
        val currentColumnIndex = Globals.userStates[chatId] ?: 0
        println("aaa3 Текущий индекс колонки: $currentColumnIndex")
        if (currentColumnIndex > 16) {
            println("aaa4 Все колонки пройдены, отправляем поздравление")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = "Поздравляем, вы завершили блок «Глаголы 1»!",
                    replyMarkup = Keyboards.finalVerbsButton()  // кнопка-заглушка
                )
            }
            return
        }

        // Получаем 7 строк (A1..A7, B1..B7 и т.д.)
        println("aaa5 Читаем 7 строк из колонки с индексом $currentColumnIndex")
        val lines = read7LinesFromColumn(filePath, sheetName, currentColumnIndex)
        println("aaa6 Прочитанные строки: $lines")
        // Формируем сообщение
        val messageText = buildVerbsMessage(lines, chatId)
        println("aaa7 Сформированное сообщение:\n$messageText")

        // Отправляем сообщение (если это последняя колонка R, то сразу выведется поздравление в следующий раз)
        val isLastColumn = (currentColumnIndex == 17)
        println("aaa8 isLastColumn = $isLastColumn")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                //parseMode = ParseMode.MARKDOWN_V2, // чтобы спойлеры корректно отображались
                replyMarkup = if (!isLastColumn) Keyboards.nextVerbsButton() else null
            )
        }
    }

    fun read7LinesFromColumn(filePath: String, sheetName: String, columnIndex: Int): List<String> {
        println("aaa9 read7LinesFromColumn // filePath=$filePath, sheetName=$sheetName, columnIndex=$columnIndex")
        val file = File(filePath)
        if (!file.exists()) {
            println("aaa10 Ошибка: Файл $filePath не найден")
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        WorkbookFactory.create(file).use { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("Лист $sheetName не найден")
            println("aaa11 Лист $sheetName найден. Всего строк: ${sheet.lastRowNum + 1}")

            val result = mutableListOf<String>()
            for (rowIndex in 0..7) {
                val row = sheet.getRow(rowIndex)
                val cell = row?.getCell(columnIndex)
                val cellValue = cell?.toString()?.trim() ?: ""
                println("aaa12 Строка ${rowIndex + 1}: \"$cellValue\"")
                result.add(cellValue)
            }
            println("aaa13 Завершено чтение строк. Результат: $result")
            return result
        }
    }

    fun buildVerbsMessage(lines: List<String>, chatId: Long): String {
        println("aaa14 buildVerbsMessage // Вход. Исходные строки: $lines")
        // 1-я строка – без изменений
        val line1 = lines[0]
        println("aaa15 line1: $line1")
        // 2-я строка – спойлер
        val line2 = lines[1]
        println("aaa16 line2 (спойлер): $line2")
        // 3–7 строки – с заменами
        val line3 = applySymbolReplacements(lines[2], chatId)
        val line4 = applySymbolReplacements(lines[3], chatId)
        val line5 = applySymbolReplacements(lines[4], chatId)
        val line6 = applySymbolReplacements(lines[5], chatId)
        val line7 = applySymbolReplacements(lines[6], chatId)
        val line8 = applySymbolReplacements(lines[7], chatId)
        println("aaa17 line3: $line3")
        println("aaa18 line4: $line4")
        println("aaa19 line5: $line5")
        println("aaa20 line6: $line6")
        println("aaa21 line7: $line7")

        val message = listOf(line1, line2, line3, line4, line5, line6, line7, line8).joinToString("\n")
        println("aaa22 Итоговое сообщение сформировано")
        return message
    }

    fun applySymbolReplacements(input: String, chatId: Long): String {
        println("aaa23 applySymbolReplacements // Вход. Исходная строка: \"$input\"")

        val sb = StringBuilder()
        val chars = input.toCharArray()
        var i = 0
        while (i < chars.size) {
            val c = chars[i]
            when (c) {
                '*' -> {
                    // Получаем уже сохранённый глагол или генерируем новый, если его ещё нет
                    val replacement = Globals.userRandomVerb.getOrPut(chatId) { getRandomVerb() }
                    println("aaaXX Замена символа *. Замена: \"$replacement\"")
                    sb.append(replacement)
                }
                '§' -> {
                    val prev = sb.lastOrNull() ?: ' '
                    val isPrevVowel = prev in listOf('a', 'i', 'u')
                    val replacement = if (isPrevVowel) "" else "i"
                    println("aaa24 Замена символа §. Предыдущий символ: '$prev', замена: \"$replacement\"")
                    sb.append(replacement)
                }
                '&' -> {
                    val prev = sb.lastOrNull() ?: ' '
                    val isSpecial = prev in listOf('p', 't', 'k', 'q', 's')
                    val replacement = if (isSpecial) "t" else "d"
                    println("aaa25 Замена символа &. Предыдущий символ: '$prev', замена: \"$replacement\"")
                    sb.append(replacement)
                }
                '%' -> {
                    val prev = sb.lastOrNull() ?: ' '
                    val isSpecial = prev in listOf('p', 't', 'k', 'q', 's')
                    val replacement = if (isSpecial) "k" else "g"
                    println("aaa26 Замена символа %. Предыдущий символ: '$prev', замена: \"$replacement\"")
                    sb.append(replacement)
                }
                '$' -> {
                    val prev = sb.lastOrNull() ?: ' '
                    val isVowel = prev in listOf('a', 'i', 'u')
                    val replacement = if (isVowel) "r" else "ar"
                    println("aaa27 Замена символа \$. Предыдущий символ: '$prev', замена: \"$replacement\"")
                    sb.append(replacement)
                }
                '^' -> {
                    val prev = sb.lastOrNull() ?: ' '
                    val isVowel = prev in listOf('a', 'i', 'u')
                    val replacement = if (isVowel) "" else "a"
                    println("aaa28 Замена символа ^. Предыдущий символ: '$prev', замена: \"$replacement\"")
                    sb.append(replacement)
                }
                '№' -> {
                    val prev = sb.lastOrNull() ?: ' '
                    val replacement = when (prev) {
                        'i' -> "b"
                        'a' -> "y"
                        else -> "a"
                    }
                    println("aaa29 Замена символа №. Предыдущий символ: '$prev', замена: \"$replacement\"")
                    sb.append(replacement)
                }
                else -> {
                    sb.append(c)
                }
            }
            i++
        }
        val result = sb.toString()
        println("aaa30 Результат после замен: \"$result\"")
        return result
    }

    fun getRandomVerb(): String {
        println("aaa_getRandomVerb // Вход в функцию получения случайного глагола")
        val filePath = "Глаголы.xlsx"
        val sheetName = "Список глаголов"
        val file = File(filePath)
        if (!file.exists()) {
            println("aaa_getRandomVerb Ошибка: Файл $filePath не найден")
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        val verbs = mutableListOf<String>()
        WorkbookFactory.create(file).use { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("Лист $sheetName не найден")
            println("aaa_getRandomVerb Лист $sheetName найден. Читаем строки...")
            // Читаем строки, начиная со второй (индекс 1) до 89-й (индекс 88)
            for (rowIndex in 1 until 89) {
                val row = sheet.getRow(rowIndex)
                if (row == null) continue
                val cell = row.getCell(0)
                val word = cell?.toString()?.trim() ?: continue
                if (word.isNotEmpty()) {
                    verbs.add(word)
                }
            }
        }
        println("aaa_getRandomVerb Список глаголов: $verbs")
        if (verbs.isEmpty()) {
            println("aaa_getRandomVerb Нет глаголов найдено!")
            throw IllegalArgumentException("Нет глаголов в листе $sheetName")
        }
        val randomVerb = verbs.random()
        println("aaa_getRandomVerb Случайный глагол: $randomVerb")
        return randomVerb
    }
}
import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFColor
import java.io.File
import kotlin.collections.joinToString
import kotlin.collections.map
import kotlin.experimental.and

object Verbs3 {

    fun callVerbs3(chatId: Long, bot: Bot) {
        println("callVerbs3: Обработка блока глаголов 3")
        Globals.userRandomVerb.remove(chatId)
        Globals.userBlocks[chatId] = 9
        Globals.userStates[chatId] = 0
        Verbs3.handleBlockVerbs3(chatId, bot)
    }

    fun handleBlockVerbs3(chatId: Long, bot: Bot) {
        println("handleBlockVerbs3: Начало обработки для chatId = $chatId")
        val currentState = Globals.userStates[chatId] ?: 0
        println("handleBlockVerbs3: Текущее состояние = $currentState")
        val ranges = Config.GVERBS_RANGES_3
        println("handleBlockVerbs3: Всего диапазонов = ${ranges.size}")
        if (currentState >= ranges.size) {
            println("handleBlockVerbs3: Текущее состояние превышает число диапазонов, отправляем сообщение об окончании")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = "Глаголы 3: все сообщения пройдены."
                )
            }
            return
        }
        val range = ranges[currentState]
        println("handleBlockVerbs3: Выбран диапазон = $range")
        // Передаем chatId для применения обработки символов
        val message = buildVerbsMessageVerbs3(Config.GVERBS_FILE, "Глаголы 3", range, chatId)
        println("handleBlockVerbs3: Сформированное сообщение:\n$message")
        // Если это не последний диапазон, добавляем кнопку "Далее"
        if (currentState < ranges.size - 1) {
            println("handleBlockVerbs3: Это не последний диапазон, добавляем кнопку 'Далее'")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = message,
                    replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                        InlineKeyboardButton.CallbackData("Далее", "next_verbs3")
                    )
                )
            }
        } else {
            println("handleBlockVerbs3: Это последний диапазон, отправляем сообщение без кнопки")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = message
                )
            }
        }
        println("handleBlockVerbs3: Завершение обработки для chatId = $chatId")
    }

    fun buildVerbsMessageVerbs3(filePath: String, sheetName: String, range: String, chatId: Long): String {
        println("buildVerbsMessageVerbs3: Начало формирования сообщения для диапазона '$range' на листе '$sheetName' из файла '$filePath'")
        val lines = readLinesFromColumnVerbs3(filePath, sheetName, range)
        println("buildVerbsMessageVerbs3: Прочитанные строки: ${lines.joinToString(separator = " | ")}")
        // Собираем строки в один текст
        val rawMessage = lines.joinToString("\n")
        println("buildVerbsMessageVerbs3: Сырой текст сообщения:\n$rawMessage")
        // Заменяем символ "–" на "-"
        val replacedMessage = rawMessage.replace("–", "-")
        println("buildVerbsMessageVerbs3: Текст после замены дефиса:\n$replacedMessage")
        // Применяем обработку символов по правилам (без экранирования Markdown)
        val finalMessage = applySymbolReplacements(replacedMessage, chatId)
        println("buildVerbsMessageVerbs3: Итоговое сообщение после обработки символов:\n$finalMessage")
        return finalMessage
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

    fun readLinesFromColumnVerbs3(filePath: String, sheetName: String, range: String): List<String> {
        println("readLinesFromColumnVerbs3: Начало обработки диапазона '$range' на листе '$sheetName' из файла '$filePath'")
        // Ожидаем формат диапазона, например: "A1-A5"
        val parts = range.split("-")
        if (parts.size != 2) {
            val errorMsg = "readLinesFromColumnVerbs3: Неверный формат диапазона: $range"
            println(errorMsg)
            throw IllegalArgumentException(errorMsg)
        }
        val startCell = parts[0]
        val endCell = parts[1]
        println("readLinesFromColumnVerbs3: startCell = $startCell, endCell = $endCell")
        // Извлекаем букву столбца и номера строк
        val startColumn = startCell.takeWhile { it.isLetter() }
        val startRow = startCell.dropWhile { it.isLetter() }.toIntOrNull()
            ?: throw IllegalArgumentException("readLinesFromColumnVerbs3: Неверный формат ячейки: $startCell")
        val endRow = endCell.dropWhile { it.isLetter() }.toIntOrNull()
            ?: throw IllegalArgumentException("readLinesFromColumnVerbs3: Неверный формат ячейки: $endCell")
        println("readLinesFromColumnVerbs3: startColumn = $startColumn, startRow = $startRow, endRow = $endRow")
        // Преобразуем букву в индекс (A -> 0, B -> 1, …)
        val columnIndex = startColumn.uppercase()[0] - 'A'
        println("readLinesFromColumnVerbs3: columnIndex = $columnIndex")
        val file = File(filePath)
        if (!file.exists()) {
            val errorMsg = "readLinesFromColumnVerbs3: Файл '$filePath' не найден"
            println(errorMsg)
            throw IllegalArgumentException(errorMsg)
        }
        println("readLinesFromColumnVerbs3: Файл '$filePath' найден")
        val excelManager = ExcelManager(filePath)
        return excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("readLinesFromColumnVerbs3: Лист '$sheetName' не найден")
            println("readLinesFromColumnVerbs3: Лист '$sheetName' найден, всего строк = ${sheet.lastRowNum + 1}")
            val lines = mutableListOf<String>()
            // Индексация строк начинается с 0, поэтому (startRow - 1) до (endRow - 1)
            for (rowIndex in (startRow - 1) until endRow) {
                val row = sheet.getRow(rowIndex)
                val cellText = row?.getCell(columnIndex)?.toString()?.trim() ?: ""
                println("readLinesFromColumnVerbs3: Строка ${rowIndex + 1}: '$cellText'")
                lines.add(cellText)
            }
            lines
        }
    }

    fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("ggg getRgbWithTint // Получение RGB цвета с учётом оттенка")

        println("g1 Входной параметр: XSSFColor=$this")
        val baseRgb = rgb
        if (baseRgb == null) {
            println("g2 ⚠️ Базовый RGB не найден.")
            return null
        }
        println("g3 Базовый RGB: ${baseRgb.joinToString { "%02X".format(it) }}")

        val tint = this.tint
        val result = if (tint != 0.0) {
            println("g4 Применяется оттенок: $tint")
            baseRgb.map { (it and 0xFF.toByte()) * (1 + tint) }
                .map { it.coerceIn(0.0, 255.0).toInt().toByte() }
                .toByteArray()
        } else {
            println("g5 Оттенок не применяется.")
            baseRgb
        }

        println("g6 Итоговый RGB с учётом оттенка: ${result?.joinToString { "%02X".format(it) }}")
        return result
    }
}
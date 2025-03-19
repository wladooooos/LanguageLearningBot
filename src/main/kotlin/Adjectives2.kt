import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.ParseMode
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.keyboards.Keyboards
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import java.io.File
import kotlin.collections.joinToString
import kotlin.collections.map
import kotlin.experimental.and

object Adjectives2 {

    fun callAdjective2(chatId: Long, bot: Bot) {
        println("callAdjective2: Выбран блок прилагательных 2")
        Globals.userBlocks[chatId] = 6
        Globals.userStates.remove(chatId)
        Globals.userPadezh.remove(chatId)
        Globals.userColumnOrder.remove(chatId)
        handleBlockAdjective2(chatId, bot)
    }

    fun callChangeWordsAdjective2(chatId: Long, bot: Bot) {
        println("callChangeWordsAdjective2: Обработка изменения набора слов для прилагательных 2")
        Globals.userReplacements.remove(chatId)
        Globals.sheetColumnPairs.remove(chatId)
        Globals.userStates.remove(chatId)
        com.github.kotlintelegrambot.dispatcher.initializeSheetColumnPairsFromFile(chatId)
        handleBlockAdjective2(chatId, bot)
    }

    // Переход к блоку с прилагательными
    fun handleBlockAdjective2(chatId: Long, bot: Bot) {
        println("UUU2 handleBlockAdjective2 // Переход к блоку с прилагательными")
        println("U21 Вход в функцию. Параметры: chatId=$chatId")

        val currentBlock = Globals.userBlocks[chatId] ?: 1
        println("U22 Текущий блок: $currentBlock, Лист: Прилагательные 2")

        if (Globals.userReplacements[chatId].isNullOrEmpty()) {
            initializeSheetColumnPairsFromFile(chatId)
            println("U23 Сгенерированы замены: ${Globals.userReplacements[chatId]}")
        }

        // Используем диапазоны для прилагательных из Config
        val rangesForAdjectives = Config.ADJECTIVE_RANGES_2
        println("U24 Диапазоны мини-блоков: $rangesForAdjectives")

        val currentState = Globals.userStates[chatId] ?: 0
        println("U25 Текущее состояние пользователя: $currentState")

        // Если все мини-блоки завершены
        if (currentState >= rangesForAdjectives.size) {
            println("U26 Все мини-блоки завершены, отправляем финальное меню.")
            sendFinalButtonsForAdjectives(chatId, bot)
            Globals.userReplacements.remove(chatId) // Очищаем замены
            return
        }

        val currentRange = rangesForAdjectives[currentState]
        println("U27 Обрабатываем текущий диапазон: $currentRange")

        // Генерация сообщения с заменами
        val messageText = try {
            generateAdjectiveMessage(Config.TABLE_FILE, "Прилагательные 2", currentRange, Globals.userReplacements[chatId]!!)
        } catch (e: Exception) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
            println("U28 Ошибка при генерации сообщения: ${e.message}")
            return
        }
        println("U29 Сообщение сгенерировано: $messageText")

        val isLastRange = currentState == rangesForAdjectives.size - 1
        println("U210 Проверка на последний диапазон: $isLastRange")

        if (Globals.userStates[chatId] == null) { // Если пользователь только вошел в блок
            sendReplacementsMessage(chatId, bot) // Отправляем сообщение с 9 парами слов
        }

        // Отправляем сообщение (либо с кнопкой "Далее", либо финальное меню)
        if (isLastRange) {
            println("Последний диапазон. Отправляем финальное меню после этого сообщения.")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    //parseMode = ParseMode.MARKDOWN_V2
                )
            }
            sendFinalButtonsForAdjectives(chatId, bot)
        } else {
            println("Не последний диапазон. Отправляем сообщение с кнопкой 'Далее'.")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    //parseMode = ParseMode.MARKDOWN_V2,
                    replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                        InlineKeyboardButton.CallbackData("Далее", "next_adjective:$chatId")
                    )
                )
            }
        }
        println("U211 Сообщение отправлено пользователю: chatId=$chatId")
    }

    // Генерация списка замен для пользователя
    fun generateReplacements(chatId: Long) {
        println("VVV generateReplacements // Генерация списка замен для пользователя $chatId")

        // Проверяем, есть ли данные для пользователя
        val userPairs = Globals.sheetColumnPairs[chatId] ?: run {
            println("❌ Ошибка: Нет данных в sheetColumnPairs для пользователя $chatId")
            return
        }

        // Берем только ключи (узбекские слова)
        val keysList = userPairs.keys.toList()
        if (keysList.size < 9) {
            println("⚠️ Предупреждение: у пользователя $chatId недостаточно пар для замены.")
        }

        // Создаем изменяемый Map (MutableMap)
        val replacements = mutableMapOf<Int, String>()
        keysList.take(9).forEachIndexed { index, key ->
            replacements[index + 1] = key
        }

        // Присваиваем userReplacements для конкретного пользователя
        Globals.userReplacements[chatId] = replacements
        println("✅ Список замен обновлен для пользователя $chatId: $replacements")
    }

    // Генерация сообщения для блока прилагательных
    fun generateAdjectiveMessage(
        filePath: String,
        sheetName: String,
        range: String,
        replacements: Map<Int, String>
    ): String {
        println("WWW generateAdjectiveMessage // Генерация сообщения для блока прилагательных")
        println("W1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName, range=$range")

        val rawText = generateMessageFromRange(filePath, sheetName, range, null, null)
        println("W2 Сырые данные из диапазона: \"$rawText\"")

        val processedText = rawText.replace(Regex("[1-9]")) { match ->
            val digit = match.value.toInt()
            replacements[digit] ?: match.value // Если нет замены, оставляем как есть
        }
        println("W3 Обработанный текст: \"$processedText\"")
        return processedText
    }

    // Отправка финального меню
    fun sendFinalButtonsForAdjectives(chatId: Long, bot: Bot) {
        println("XXX sendFinalButtonsForAdjectives // Отправка финального меню")
        println("X1 Вход в функцию. Параметры: chatId=$chatId")

        val currentBlock = Globals.userBlocks[chatId] ?: 5  // По умолчанию 5 (прилагательные 1)
        val repeatCallback = if (currentBlock == 5) "block:adjective1" else "block:adjective2"
        val changeWordsCallback = if (currentBlock == 5) "change_words_adjective1" else "change_words_adjective2"

        val navigationButton = if (currentBlock == 5) {
            InlineKeyboardButton.CallbackData("Следующий блок", "block:adjective2")
        } else {
            InlineKeyboardButton.CallbackData("Предыдущий блок", "block:adjective1")
        }

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Вы завершили все этапы работы с этим блоком прилагательных. Что будем делать дальше?",
                replyMarkup = Keyboards.finalAdjectiveButtons(Globals.userBlocks[chatId] ?: 5)
            )
        }
        println("X2 Финальное меню отправлено пользователю: chatId=$chatId")
    }

    // Генерация сообщения из диапазона Excel
    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?): String {
        println("SSS generateMessageFromRange // Генерация сообщения из диапазона Excel")
        println("S1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName, range=$range, wordUz=$wordUz, wordRus=$wordRus")

        val file = File(filePath)
        if (!file.exists()) {
            println("S2 Ошибка: файл $filePath не найден.")
            throw IllegalArgumentException("Файл $filePath не найден")
        }

        val excelManager = ExcelManager(filePath)
        val result = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                println("S3 Ошибка: лист $sheetName не найден.")
                throw IllegalArgumentException("Лист $sheetName не найден")
            }

            println("S4 Лист $sheetName найден. Извлекаем ячейки из диапазона $range")
            val cells = extractCellsFromRange(sheet, range, wordUz)
            println("S5 Извлеченные ячейки: $cells")

            val firstCell = cells.firstOrNull() ?: ""
            val messageBody = cells.drop(1).joinToString("\n")

            val res = listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
            println("S7 Генерация завершена. Результат: $res")
            res
        }
        println("S6 Файл Excel закрыт.")
        return result
    }

    // Экранирование Markdown V2
    fun String.escapeMarkdownV2(): String {
        println("TTT escapeMarkdownV2 // Экранирование Markdown V2")
        println("T1 Вход в функцию. Строка: \"$this\"")
        val escaped = this.replace("\\", "\\\\")
            .replace("_", "\\_")
            .replace("*", "\\*")
            .replace("[", "\\[")
            .replace("]", "\\]")
            .replace("(", "\\(")
            .replace(")", "\\)")
            .replace("~", "\\~")
            .replace("`", "\\`")
            .replace(">", "\\>")
            .replace("#", "\\#")
            .replace("+", "\\+")
            .replace("-", "\\-")
            .replace("=", "\\=")
            .replace("|", "\\|")
            .replace("{", "\\{")
            .replace("}", "\\}")
            .replace(".", "\\.")
            .replace("!", "\\!")
        println("T2 Экранирование завершено. Результат: \"$escaped\"")
        return escaped
    }

    // Обработка узбекского слова в контексте строки
    fun adjustWordUz(content: String, wordUz: String?): String {
        println("UUU adjustWordUz // Обработка узбекского слова в контексте строки")
        println("U1 Вход в функцию. Параметры: content=\"$content\", wordUz=\"$wordUz\"")

        fun Char.isVowel() = this.lowercaseChar() in "aeiouаеёиоуыэюя"

        val result = buildString {
            var i = 0
            while (i < content.length) {
                val char = content[i]
                when {
                    char == '+' && i + 1 < content.length -> {
                        val nextChar = content[i + 1]
                        val lastChar = wordUz?.lastOrNull()
                        val replacement = when {
                            lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> "s"
                            lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> "i"
                            else -> ""
                        }
                        append(replacement)
                        append(nextChar)
                        i++
                    }
                    char == '*' -> append(wordUz)
                    else -> append(char)
                }
                i++
            }
        }
        println("U2 Результат обработки: \"$result\"")
        return result
    }

    fun getRangeIndices(range: String): Pair<Int, Int> {
        val parts = range.split("-")
        if (parts.size != 2) {
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        return start to end
    }

    fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        val cell = row.getCell(column)
        return cell?.let { processCellContent(it, wordUz) }
    }

    fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("VVV extractCellsFromRange // Извлечение и обработка ячеек из диапазона")
        println("V1 Вход в функцию. Параметры: range=\"$range\", wordUz=\"$wordUz\"")

        val (start, end) = getRangeIndices(range)
        val column = range[0] - 'A'
        println("V2 Диапазон строк: $start-$end, Колонка: $column")

        return (start..end).mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex)
            if (row == null) {
                println("V3 ⚠️ Строка $rowIndex отсутствует, пропускаем.")
                null
            } else {
                val processed = processRowForRange(row, column, wordUz)
                println("V5 ✅ Обработанная ячейка в строке $rowIndex: \"$processed\"")
                processed
            }
        }.also { cells ->
            println("V6 Результат извлечения: $cells")
        }
    }

    // Обработка содержимого ячейки
    fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("bbb processCellContent // Обработка содержимого ячейки")

        println("b1 Входные параметры: cell=$cell, wordUz=$wordUz")
        if (cell == null) {
            println("b2 Ячейка пуста. Возвращаем пустую строку.")
            return ""
        }

        val richText = cell.richStringCellValue as XSSFRichTextString
        val text = richText.string
        println("b3 Извлечённое содержимое ячейки: \"$text\"")

        val runs = richText.numFormattingRuns()
        println("b4 Количество форматированных участков текста: $runs")

        val result = if (runs == 0) {
            println("b5 Нет форматированных участков. Переходим к обработке без форматирования.")
            processCellWithoutRuns(cell, text, wordUz)
        } else {
            println("b6 Есть форматированные участки. Переходим к обработке форматированных участков.")
            processFormattedRuns(richText, text, wordUz)
        }

        println("b7 Результат обработки ячейки: \"$result\"")
        return result
    }

    // Обработка ячейки без форматированных участков
    fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("ccc processCellWithoutRuns // Обработка ячейки без форматированных участков")

        println("c1 Входные параметры: cell=$cell, text=$text, wordUz=$wordUz")
        val font = getCellFont(cell)
        println("c2 Полученный шрифт ячейки: $font")

        val isRed = font != null && getFontColor(font) == "#FF0000"
        println("c3 Цвет текста красный: $isRed")

        val result = if (isRed) {
            println("c4 Вся ячейка имеет красный цвет. Блюрим текст.")
            "||${adjustWordUz(text, wordUz).escapeMarkdownV2()}||"
        } else {
            println("c5 Текст не красный. Оставляем текст без изменений.")
            adjustWordUz(text, wordUz).escapeMarkdownV2()
        }

        println("c6 Результат обработки текста: \"$result\"")
        return result
    }

    // Обработка форматированных участков текста
    fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("ddd processFormattedRuns // Обработка форматированных участков текста")

        println("d1 Входные параметры: richText=$richText, text=$text, wordUz=$wordUz")
        val result = buildString {
            for (i in 0 until richText.numFormattingRuns()) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < richText.numFormattingRuns()) richText.getIndexOfFormattingRun(i + 1) else text.length
                val substring = text.substring(start, end)

                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
                println("d2 🎨 Цвет участка $i: $colorHex")

                val adjustedSubstring = adjustWordUz(substring, wordUz)

                if (colorHex == "#FF0000") {
                    println("d3 🔴 Текст участка \"$substring\" красный. Добавляем блюр.")
                    append("||${adjustedSubstring.escapeMarkdownV2()}||")
                } else {
                    println("d4 Текст участка \"$substring\" не красный. Оставляем как есть.")
                    append(adjustedSubstring.escapeMarkdownV2())
                }
            }
        }
        println("d5 ✅ Результат обработки форматированных участков: \"$result\"")
        return result
    }

    // Получение шрифта ячейки
    fun getCellFont(cell: Cell): XSSFFont? {
        println("eee getCellFont // Получение шрифта ячейки")

        println("e1 Входной параметр: cell=$cell")
        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
        if (workbook == null) {
            println("e2 ❌ Ошибка: Невозможно получить шрифт, workbook не является XSSFWorkbook.")
            return null
        }
        val fontIndex = cell.cellStyle.fontIndexAsInt
        println("e3 Индекс шрифта: $fontIndex")

        val font = workbook.getFontAt(fontIndex) as? XSSFFont
        println("e4 Результат: font=$font")
        return font
    }

    // Функция для извлечения цвета шрифта
    fun getFontColor(font: XSSFFont): String {
        println("fff getFontColor // Извлечение цвета шрифта")

        println("f1 Входной параметр: font=$font")
        val xssfColor = font.xssfColor
        if (xssfColor == null) {
            println("f2 ⚠️ Цвет шрифта не определён.")
            return "Цвет не определён"
        }

        val rgb = xssfColor.rgb
        val result = if (rgb != null) {
            val colorHex = rgb.joinToString(prefix = "#", separator = "") { "%02X".format(it) }
            println("f3 🎨 Цвет шрифта в формате HEX: $colorHex")
            colorHex
        } else {
            println("f4 ⚠️ RGB не найден.")
            "Цвет не определён"
        }

        println("f5 Результат: $result")
        return result
    }

    // Вспомогательный метод для обработки цветов с учётом оттенков
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

    // Отправка сообщения с 9 парами слов из sheetColumnPairs
    fun sendReplacementsMessage(chatId: Long, bot: Bot) {
        println("### sendReplacementsMessage // Отправка сообщения с 9 парами слов из sheetColumnPairs")

        val userPairs = Globals.sheetColumnPairs[chatId]

        if (userPairs.isNullOrEmpty()) {
            println("❌ Ошибка: Данные в sheetColumnPairs отсутствуют для пользователя $chatId.")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = "Ошибка: Данные для отправки не найдены."
                )
            }
            return
        }

        // Формируем сообщение в нужном формате
        val messageText = userPairs.entries.joinToString("\n") { (key, value) ->
            "$key - $value"
        }
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText
            )
        }

        println("✅ Сообщение отправлено пользователю $chatId:\n$messageText")
    }

    // Инициализация пар для пользователя
    fun initializeSheetColumnPairsFromFile(chatId: Long) {
        println("### initializeSheetColumnPairsFromFile // Инициализация пар для пользователя $chatId")

        Globals.sheetColumnPairs[chatId] = mutableMapOf()  // Создаем пустой мап для пользователя
        val file = File(Config.TABLE_FILE)
        if (!file.exists()) {
            println("❌ Файл $Config.TABLE_FILE не найден!")
            return
        }

        val excelManager = ExcelManager(Config.TABLE_FILE)
        val sheetNames = listOf("Существительные", "Глаголы", "Прилагательные")
        val userPairs = mutableMapOf<String, String>()  // Временный мап для хранения пар

        excelManager.useWorkbook { workbook ->
            for (sheetName in sheetNames) {
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    println("⚠️ Лист $sheetName не найден!")
                    continue  // ✅ Просто continue, без run {}
                }

                val candidates = mutableListOf<Pair<String, String>>() // Кандидаты на выборку

                for (i in 0..sheet.lastRowNum) {
                    val row = sheet.getRow(i) ?: continue
                    val key = row.getCell(0)?.toString()?.trim() ?: ""
                    val value = row.getCell(1)?.toString()?.trim() ?: ""
                    if (key.isNotEmpty() && value.isNotEmpty()) {
                        candidates.add(key to value)
                    }
                }

                if (candidates.size < 3) {
                    println("⚠️ Недостаточно данных в листе $sheetName для выбора 3 пар.")
                    continue
                }

                val selectedPairs = candidates.shuffled().take(3)  // Берем 3 случайные пары
                for ((key, value) in selectedPairs) {
                    userPairs[key] = value
                }
            }
        }

        if (userPairs.size == 9) {
            Globals.sheetColumnPairs[chatId] = userPairs  // Сохраняем пары в общий мап
            println("✅ Успешно загружены 9 пар для пользователя $chatId: $userPairs")
        } else {
            println("❌ Ошибка: Получено ${userPairs.size} пар вместо 9.")
        }
        generateReplacements(chatId)
    }
}
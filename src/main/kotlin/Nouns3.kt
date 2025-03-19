import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.initializeUserBlockStates
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.ParseMode
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.keyboards.Keyboards
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import java.io.File
import kotlin.collections.joinToString
import kotlin.collections.map
import kotlin.experimental.and

object Nouns3 {

    fun callNouns3(chatId: Long, bot: Bot) {
        println("callNouns3: Инициализация состояний для блока 3")
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val (block1Completed, block2Completed, _) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
        if (block1Completed && block2Completed) {
            Globals.userBlocks[chatId] = 3
            handleBlock3(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
        } else {
            val notCompletedBlocks = mutableListOf<String>()
            if (!block1Completed) notCompletedBlocks.add("Блок 1")
            if (!block2Completed) notCompletedBlocks.add("Блок 2")
            val messageText = "Вы не завершили следующие блоки:\n" +
                    notCompletedBlocks.joinToString("\n") + "\nПройдите их перед переходом к 3-му блоку."
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                        InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
                    )
                )
            }
        }
    }

    fun handleBlock3(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("🚀 handleBlock3 // Запуск блока 3 для пользователя $chatId")

        if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
            sendWordMessage(chatId, bot, filePath)
            println("❌ Ошибка: слова не выбраны, запрошен повторный выбор.")
            return
        }

        // Используем неизменяемый список диапазонов из Config
        val allRanges = Config.ALL_RANGES_BLOCK_3

        // Если у пользователя еще нет порядка столбцов — перемешиваем
        if (Globals.userColumnOrder[chatId].isNullOrEmpty()) {
            Globals.userColumnOrder[chatId] = allRanges.shuffled().toMutableList()
            println("🔀 Перемешали мини-блоки для пользователя $chatId: ${Globals.userColumnOrder[chatId]}")
        }

        val shuffledRanges = Globals.userColumnOrder[chatId]!!
        val currentState = Globals.userStates[chatId] ?: 0
        println("📌 Текущее состояние пользователя: $currentState")

        // Загружаем баллы пользователя
        val completedRanges = getCompletedRanges(chatId, filePath)
        println("✅ Пройденные мини-блоки: $completedRanges")

        // Ищем первый неповторяющийся диапазон
        var currentRange: String? = null
        for (range in shuffledRanges) {
            if (!completedRanges.contains(range)) {
                currentRange = range
                break
            }
        }

        // Если ВСЕ 30 диапазонов пройдены, то берем текущий диапазон из списка
        if (currentRange == null) {
            println("⚠️ ВСЕ 30 мини-блоков уже пройдены! Берем текущий из списка.")
            currentRange = shuffledRanges[currentState % shuffledRanges.size]
        }

        println("🎯 Выбран диапазон: $currentRange")

        val messageText = try {
            generateMessageFromRange(filePath, "Существительные 3", currentRange!!, wordUz, wordRus)
        } catch (e: Exception) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
            println("❌ Ошибка при генерации сообщения: ${e.message}")
            return
        }

        val isLastRange = currentState == 5 // 6 сообщений всего
        println("📩 Отправка сообщения: \"$messageText\"")

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                //parseMode = ParseMode.MARKDOWN_V2,
                replyMarkup = if (isLastRange) null else Keyboards.nextButton(wordUz, wordRus)
            )
        }

        // Сохраняем прогресс
        saveUserProgressBlok3(chatId, filePath, currentRange!!)
        println("💾 Прогресс сохранен: $currentRange")

        if (!isLastRange) {
            println("📌 Переход к следующему шагу")
        } else {
            println("🏁 Мини-блоки завершены, вызываем финальное меню")
            updateUserProgressForMiniBlocks(chatId, filePath, shuffledRanges.indices.toList())
            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        }
    }

    // Отправка клавиатуры с выбором слов
    fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
        println("KKK sendWordMessage // Отправка клавиатуры с выбором слов")
        println("K1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath")

        if (!File(filePath).exists()) {
            println("K2 Ошибка: файл $filePath не найден.")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = "Ошибка: файл с данными не найден."
                )
            }
            println("K3 Сообщение об ошибке отправлено пользователю $chatId")
            return
        }

        val inlineKeyboard = try {
            println("K4 Генерация клавиатуры из файла $filePath")
            createWordSelectionKeyboardFromExcel(filePath, "Существительные")
        } catch (e: Exception) {
            println("K5 Ошибка при обработке данных: ${e.message}")
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = "Ошибка при обработке данных: ${e.message}"
                )
            }
            println("K6 Сообщение об ошибке отправлено пользователю $chatId")
            return
        }

        println("K7 Клавиатура успешно создана: $inlineKeyboard")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Выберите слово из списка:",
                replyMarkup = inlineKeyboard
            )
        }
        println("K8 Сообщение с клавиатурой отправлено пользователю $chatId")
    }

    // Создание клавиатуры из Excel-файла
    fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
        println("LLL createWordSelectionKeyboardFromExcel // Создание клавиатуры из Excel-файла")
        println("L1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName")

        println("L2 Проверка существования файла $filePath")
        val file = File(filePath)
        if (!file.exists()) {
            println("L3 Ошибка: файл $filePath не найден")
            throw IllegalArgumentException("Файл $filePath не найден")
        }

        val excelManager = ExcelManager(filePath)
        val buttons = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("D5 Ошибка: Лист $sheetName не найден")
            println("L6 Лист найден: $sheetName")
            println("L7 Всего строк в листе: ${sheet.lastRowNum + 1}")
            val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
            println("L8 Случайные строки для кнопок: $randomRows")
            println("L9 Начало обработки строк для создания кнопок")
            val buttons = randomRows.mapNotNull { rowIndex ->
                val row = sheet.getRow(rowIndex)
                if (row == null) {
                    println("L10 Строка $rowIndex отсутствует, пропускаем")
                    return@mapNotNull null
                }

                val wordUz = row.getCell(0)?.toString()?.trim()
                val wordRus = row.getCell(1)?.toString()?.trim()

                if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
                    println("L11 Некорректные данные в строке $rowIndex: wordUz=$wordUz, wordRus=$wordRus. Пропускаем")
                    return@mapNotNull null
                }

                println("L12 Обработана строка $rowIndex: wordUz=$wordUz, wordRus=$wordRus")
                InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
            }.chunked(2)
            println("L13 Кнопки успешно созданы. Количество строк кнопок: ${buttons.size}")
            buttons
        }
        println("L14 Файл Excel закрыт. Генерация клавиатуры завершена")

        return InlineKeyboardMarkup.create(buttons)
    }

    // Отправка финального меню действий
    fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
        println("RRR sendFinalButtons // Отправка финального меню действий")
        println("R1 Вход в функцию. Параметры: chatId=$chatId, wordUz=$wordUz, wordRus=$wordRus")

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
                replyMarkup = Keyboards.finalButtons(wordUz, wordRus, Globals.userBlocks[chatId] ?: 1)
            )
        }
        println("R2 Финальное меню отправлено пользователю $chatId")
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

    // Новая функция: Обновление прогресса пользователя для мини-блоков
    fun updateUserProgressForMiniBlocks(chatId: Long, filePath: String, completedMiniBlocks: List<Int>) {
        println("iii updateUserProgressForMiniBlocks // Обновление прогресса пользователя по мини-блокам")
        println("i1 Входные параметры: chatId=$chatId, filePath=$filePath, completedMiniBlocks=$completedMiniBlocks")

        val file = File(filePath)
        if (!file.exists()) {
            println("i2 Ошибка: Файл $filePath не найден.")
            throw IllegalArgumentException("Файл $filePath не найден.")
        }

        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя")
                ?: throw IllegalArgumentException("Лист 'Состояние пользователя' не найден.")

            // Ищем строку пользователя
            for (row in sheet) {
                val idCell = row.getCell(0)
                val chatIdFromCell = when (idCell?.cellType) {
                    CellType.NUMERIC -> idCell.numericCellValue.toLong()
                    CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                    else -> null
                }

                if (chatIdFromCell == chatId) {
                    println("i4 Пользователь найден. Обновляем прогресс.")
                    // Обновляем ячейки от L (11) и дальше
                    completedMiniBlocks.forEach { miniBlock ->
                        val columnIndex = 11 + miniBlock // L = 11, M = 12 и т.д.
                        val cell = row.getCell(columnIndex) ?: row.createCell(columnIndex)
                        val currentValue = cell.numericCellValue.takeIf { it > 0 } ?: 0.0
                        cell.setCellValue(currentValue + 1)
                        println("i5 Мини-блок $miniBlock: старое значение = $currentValue, новое значение = ${currentValue + 1}")
                    }
                    excelManager.safelySaveWorkbook(workbook)
                    println("i6 Прогресс успешно обновлен для пользователя $chatId.")
                    return@useWorkbook
                }
            }
            println("i7 Ошибка: Пользователь $chatId не найден в таблице.")
        }
    }

    // Сохранение прогресса пользователя
    fun saveUserProgressBlok3(chatId: Long, filePath: String, range: String) {
        println("📌 saveUserProgressBlok3 // Сохранение прогресса пользователя")

        val columnIndex = Config.COLUMN_MAPPING[range]
        if (columnIndex == null) {
            println("❌ Ошибка: Диапазон $range не найден в карте соответствий.")
            return
        }
        println("✅ Диапазон $range -> Записываем в колонку $columnIndex")

        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя 3 блок") ?: workbook.createSheet("Состояние пользователя 3 блок")
            val headerRow = sheet.getRow(0) ?: sheet.createRow(0)
            var userColumnIndex = -1

            for (col in 0 until 31) { // Проверяем до 30 столбца
                val cell = headerRow.getCell(col) ?: headerRow.createCell(col)
                if (cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                    userColumnIndex = col
                    break
                } else if (cell.cellType == CellType.BLANK) { // Если нашли пустой, записываем ID
                    cell.setCellValue(chatId.toDouble())
                    userColumnIndex = col
                    break
                }
            }

            if (userColumnIndex == -1) {
                println("⚠️ Нет свободного места для записи ID!")
                return@useWorkbook
            }

            // Получаем строку для записи балла (по индексу столбца)
            val row = sheet.getRow(columnIndex) ?: sheet.createRow(columnIndex)
            val cell = row.getCell(userColumnIndex) ?: row.createCell(userColumnIndex)
            val currentScore = cell.numericCellValue.takeIf { it > 0 } ?: 0.0
            cell.setCellValue(currentScore + 1)

            println("✅ Прогресс обновлён: chatId=$chatId, столбец=$userColumnIndex, строка=$columnIndex, новый балл=${currentScore + 1}")

            excelManager.safelySaveWorkbook(workbook)
        }
    }

    // Проверка пройденных мини-блоков
    fun getCompletedRanges(chatId: Long, filePath: String): Set<String> {
        println("📊 getCompletedRanges // Проверка пройденных мини-блоков")

        val completedRanges = mutableSetOf<String>()
        val file = File(filePath)
        if (!file.exists()) {
            println("❌ Ошибка: Файл $filePath не найден.")
            return emptySet()
        }

        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя 3 блок") ?: return@useWorkbook
            val headerRow = sheet.getRow(0) ?: return@useWorkbook

            var userColumnIndex: Int? = null
            for (col in 0 until 31) {
                val cell = headerRow.getCell(col)
                if (cell != null && cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                    userColumnIndex = col
                    break
                }
            }

            if (userColumnIndex == null) {
                println("⚠️ Пользователь $chatId не найден в таблице прогресса.")
                return@useWorkbook
            }

            val rangeMapping = listOf(
                "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7",
                "A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14",
                "A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21",
                "A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28",
                "A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"
            )

            for ((index, range) in rangeMapping.withIndex()) {
                val row = sheet.getRow(index + 1) ?: continue
                val cell = row.getCell(userColumnIndex)
                val value = cell?.numericCellValue ?: 0.0
                if (value > 0) {
                    completedRanges.add(range)
                }
            }
        }

        println("✅ Завершенные диапазоны для пользователя $chatId: $completedRanges")
        return completedRanges
    }
}
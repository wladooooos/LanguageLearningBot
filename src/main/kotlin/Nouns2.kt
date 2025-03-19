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

object Nouns2 {
    fun callNouns2(chatId: Long, bot: Bot) {
        println("Nouns2 callNouns2: Инициализация состояний для блока 2")
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val (block1Completed, _, _) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
        if (block1Completed) {
            Globals.userBlocks[chatId] = 2
            handleBlock2(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
        } else {
            val messageText = "Вы не завершили Блок 1.\nПройдите его перед переходом ко 2-му блоку."
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
    // Добавляем инициализацию состояний блоков
    fun handleBlock2(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("Nouns2 handleBlock2 // Обработка блока 2")
        // Если не выбран падеж — отправляем меню выбора
        if (Globals.userPadezh[chatId] == null) {
            sendPadezhSelection(chatId, bot, filePath)
            return
        }
        // Если не выбрано слово — отправляем меню выбора
        if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
            sendWordMessage(chatId, bot, filePath)
            return
        }
        // Определяем диапазоны для выбранного падежа из Config
        val blockRanges = Config.PADEZH_RANGES[Globals.userPadezh[chatId]] ?: return
        // Если порядок столбцов не сгенерирован — перемешиваем
        if (Globals.userColumnOrder[chatId].isNullOrEmpty()) {
            Globals.userColumnOrder[chatId] = blockRanges.shuffled().toMutableList()
        }
        val currentState = Globals.userStates[chatId] ?: 0
        val shuffledColumns = Globals.userColumnOrder[chatId]!!
        // Если ВСЕ столбцы пройдены — отправляем финальное сообщение
        if (currentState >= shuffledColumns.size) {
            addScoreForPadezh(chatId, Globals.userPadezh[chatId].toString(), filePath, block = 2)
            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
            return
        }
        val range = shuffledColumns[currentState]
        val messageText = generateMessageFromRange(filePath, "Существительные 2", range, wordUz, wordRus)
        val isLastMessage = currentState == shuffledColumns.size - 1
        // Отправляем сообщение
        if (isLastMessage) {
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    //parseMode = ParseMode.MARKDOWN_V2
                )
            }
            addScoreForPadezh(chatId, Globals.userPadezh[chatId].toString(), filePath, block = 2)
            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        } else {
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    //parseMode = ParseMode.MARKDOWN_V2,
                    replyMarkup = Keyboards.nextButton(wordUz, wordRus)
                )
            }
        }
    }

    // Формирование клавиатуры для выбора падежа
    fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
        println("Nouns2 sendPadezhSelection // Формирование клавиатуры для выбора падежа")

        Globals.userPadezh.remove(chatId)
        Globals.userStates.remove(chatId)

        val currentBlock = Globals.userBlocks[chatId] ?: 1

        val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
        if (PadezhColumns == null) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: невозможно загрузить данные блока.")
            return
        }

        val userScores = getUserScoresForBlock(chatId, filePath, PadezhColumns)

        val buttons = generatePadezhSelectionButtons(currentBlock, PadezhColumns, userScores)

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Выберите падеж для изучения блока $currentBlock:",
                replyMarkup = InlineKeyboardMarkup.create(buttons)
            )
        }
    }

    // Получение колонок текущего блока
    fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
        println("Nouns2 getPadezhColumnsForBlock // Получение колонок текущего блока")
        val result = Config.COLUMN_RANGES[block]
        return result
    }

    // Чтение баллов пользователя для блока
    fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
        println("Nouns2 getUserScoresForBlock // Чтение баллов пользователя для блока")

        val file = File(filePath)
        if (!file.exists()) {
            return emptyMap()
        }

        val scores = mutableMapOf<String, Int>()
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя")
            if (sheet == null) {
                return@useWorkbook
            }

            for (row in sheet) {
                val idCell = row.getCell(0)
                val chatIdXlsx = when (idCell?.cellType) {
                    CellType.NUMERIC -> idCell.numericCellValue.toLong()
                    CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                    else -> null
                }

                if (chatId == chatIdXlsx) {
                    for ((PadezhName, colIndex) in PadezhColumns) {
                        val cell = row.getCell(colIndex)
                        val score = cell?.numericCellValue?.toInt() ?: 0
                        scores[PadezhName] = score
                        println("G6 Падеж: $PadezhName, Баллы: $score")
                    }
                    break
                }
            }
        }
        return scores
    }

    // Формирование кнопок выбора падежей
    fun generatePadezhSelectionButtons(
        currentBlock: Int,
        PadezhColumns: Map<String, Int>,
        userScores: Map<String, Int>
    ): List<List<InlineKeyboardButton>> {
        println("Nouns2 generatePadezhSelectionButtons // Формирование кнопок выбора падежей")

        val buttons = PadezhColumns.keys.map { PadezhName ->
            val score = userScores[PadezhName] ?: 0
            InlineKeyboardButton.CallbackData("$PadezhName [$score]", "Padezh:$PadezhName")
        }.map { listOf(it) }.toMutableList()
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")))
//        if (currentBlock > 1) {
//            buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
//        }
//        if (currentBlock < 3) {
//            buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
//        }
        return buttons
    }

    // Отправка клавиатуры с выбором слов
    fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
        println("Nouns2 sendWordMessage // Отправка клавиатуры с выбором слов")

        if (!File(filePath).exists()) {
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка: файл с данными не найден."
            )
            return
        }

        val inlineKeyboard = try {
            createWordSelectionKeyboardFromExcel(filePath, "Существительные")
        } catch (e: Exception) {
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка при обработке данных: ${e.message}"
            )
            return
        }
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Выберите слово из списка:",
                replyMarkup = inlineKeyboard
            )
        }
    }

    // Создание клавиатуры из Excel-файла
    fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
        println("Nouns2 createWordSelectionKeyboardFromExcel // Создание клавиатуры из Excel-файла")
        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }

        val excelManager = ExcelManager(filePath)
        val buttons = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("D5 Ошибка: Лист $sheetName не найден")
            val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
            val buttons = randomRows.mapNotNull { rowIndex ->
                val row = sheet.getRow(rowIndex)
                if (row == null) {
                    return@mapNotNull null
                }

                val wordUz = row.getCell(0)?.toString()?.trim()
                val wordRus = row.getCell(1)?.toString()?.trim()

                if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
                    return@mapNotNull null
                }
                InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
            }.chunked(2)
            buttons
        }
        return InlineKeyboardMarkup.create(buttons)
    }

    // Отправка финального меню действий
    fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
        println("Nouns2 sendFinalButtons // Отправка финального меню действий")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
                replyMarkup = Keyboards.finalButtons(wordUz, wordRus, Globals.userBlocks[chatId] ?: 1)
            )
        }
    }

    // Генерация сообщения из диапазона Excel
    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?): String {
        println("Nouns2 generateMessageFromRange // Генерация сообщения из диапазона Excel")
        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }

        val excelManager = ExcelManager(filePath)
        val result = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                throw IllegalArgumentException("Лист $sheetName не найден")
            }
            val cells = extractCellsFromRange(sheet, range, wordUz)
            val firstCell = cells.firstOrNull() ?: ""
            val messageBody = cells.drop(1).joinToString("\n")
            val res = listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
            res
        }
        return result
    }

    // Экранирование Markdown V2
    fun String.escapeMarkdownV2(): String {
        println("Nouns2 escapeMarkdownV2 // Экранирование Markdown V2")
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
        return escaped
    }

    // Обработка узбекского слова в контексте строки
    fun adjustWordUz(content: String, wordUz: String?): String {
        println("Nouns2 adjustWordUz // Обработка узбекского слова в контексте строки")
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
        return result
    }

    fun getRangeIndices(range: String): Pair<Int, Int> {
        println("Nouns2 getRangeIndices")
        val parts = range.split("-")
        if (parts.size != 2) {
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        return start to end
    }

    fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        println("Nouns2 processRowForRang")
        val cell = row.getCell(column)
        return cell?.let { processCellContent(it, wordUz) }
    }

    fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("Nouns2 extractCellsFromRange // Извлечение и обработка ячеек из диапазона")

        val (start, end) = getRangeIndices(range)
        val column = range[0] - 'A'

        return (start..end).mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex)
            if (row == null) {
                null
            } else {
                val processed = processRowForRange(row, column, wordUz)
                processed
            }
        }.also { cells ->
        }
    }

    // Начало добавления балла для пользователя
    fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, block: Int) {
        println("Nouns2 addScoreForPadezh // Начало добавления балла для пользователя")

        val columnRanges = mapOf(
            1 to mapOf("Именительный" to 1, "Родительный" to 2, "Винительный" to 3, "Дательный" to 4, "Местный" to 5, "Исходный" to 6),
            2 to mapOf("Именительный" to 7, "Родительный" to 8, "Винительный" to 9, "Дательный" to 10, "Местный" to 11, "Исходный" to 12),
            3 to mapOf("Именительный" to 13, "Родительный" to 14, "Винительный" to 15, "Дательный" to 16, "Местный" to 17, "Исходный" to 18)
        )

        val column = columnRanges[block]?.get(Padezh)
        if (column == null) {
            println("❌ Ошибка: колонка для блока $block и падежа $Padezh не найдена.")
            return
        }

        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя")
                ?: throw IllegalArgumentException("A3 Ошибка: Лист 'Состояние пользователя' не найден")
            var userFound = false
            for (row in sheet) {
                val idCell = row.getCell(0)
                if (idCell == null) {
                    println("⚠️ Ячейка ID отсутствует в строке ${row.rowNum + 1}, пропускаем.")
                    continue
                }
                val idFromCell = try {
                    idCell.numericCellValue.toLong()
                } catch (e: Exception) {
                    println("❌ Ошибка: не удалось прочитать ID из ячейки в строке ${row.rowNum + 1}. Значение: $idCell")
                    continue
                }
                if (idFromCell == chatId) {
                    userFound = true
                    val targetCell = row.getCell(column) ?: row.createCell(column)
                    val currentValue = targetCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                    targetCell.setCellValue(currentValue + 1)
                    excelManager.safelySaveWorkbook(workbook)
                    return@useWorkbook
                }
            }
            if (!userFound) {
                println("⚠️ Пользователь с ID $chatId не найден. Новая запись не создана.")
            }
        }
    }

    // Обработка содержимого ячейки
    fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("Nouns2 processCellContent // Обработка содержимого ячейки")

        if (cell == null) {
            return ""
        }
        val richText = cell.richStringCellValue as XSSFRichTextString
        val text = richText.string
        val runs = richText.numFormattingRuns()
        val result = if (runs == 0) {
            processCellWithoutRuns(cell, text, wordUz)
        } else {
            processFormattedRuns(richText, text, wordUz)
        }
        return result
    }

    // Обработка ячейки без форматированных участков
    fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("Nouns2 processCellWithoutRuns // Обработка ячейки без форматированных участков")
        val font = getCellFont(cell)
        val isRed = font != null && getFontColor(font) == "#FF0000"
        val result = if (isRed) {
            "||${adjustWordUz(text, wordUz).escapeMarkdownV2()}||"
        } else {
            adjustWordUz(text, wordUz).escapeMarkdownV2()
        }
        return result
    }

    // Обработка форматированных участков текста
    fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("Nouns2 processFormattedRuns // Обработка форматированных участков текста")
        val result = buildString {
            for (i in 0 until richText.numFormattingRuns()) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < richText.numFormattingRuns()) richText.getIndexOfFormattingRun(i + 1) else text.length
                val substring = text.substring(start, end)
                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
                val adjustedSubstring = adjustWordUz(substring, wordUz)
                if (colorHex == "#FF0000") {
                    append("||${adjustedSubstring.escapeMarkdownV2()}||")
                } else {
                    append(adjustedSubstring.escapeMarkdownV2())
                }
            }
        }
        return result
    }

    // Получение шрифта ячейки
    fun getCellFont(cell: Cell): XSSFFont? {
        println("Nouns2 getCellFont // Получение шрифта ячейки")
        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
        if (workbook == null) {
            println("e2 ❌ Ошибка: Невозможно получить шрифт, workbook не является XSSFWorkbook.")
            return null
        }
        val fontIndex = cell.cellStyle.fontIndexAsInt
        val font = workbook.getFontAt(fontIndex) as? XSSFFont
        return font
    }

    // Функция для извлечения цвета шрифта
    fun getFontColor(font: XSSFFont): String {
        println("Nouns2 getFontColor // Извлечение цвета шрифта")
        val xssfColor = font.xssfColor
        if (xssfColor == null) {
            return "Цвет не определён"
        }

        val rgb = xssfColor.rgb
        val result = if (rgb != null) {
            val colorHex = rgb.joinToString(prefix = "#", separator = "") { "%02X".format(it) }
            colorHex
        } else {
            "Цвет не определён"
        }
        return result
    }

    // Вспомогательный метод для обработки цветов с учётом оттенков
    fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("Nouns2 getRgbWithTint // Получение RGB цвета с учётом оттенка")
        val baseRgb = rgb
        if (baseRgb == null) {
            return null
        }
        val tint = this.tint
        val result = if (tint != 0.0) {
            baseRgb.map { (it and 0xFF.toByte()) * (1 + tint) }
                .map { it.coerceIn(0.0, 255.0).toInt().toByte() }
                .toByteArray()
        } else {
            baseRgb
        }
        return result
    }
}
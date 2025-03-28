import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.initializeUserBlockStates
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
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
import kotlin.collections.mapNotNull
import kotlin.experimental.and

object Nouns2 {
    fun callNouns2(chatId: Long, bot: Bot, callbackQueryId: String) {
        println("Nouns2 callNouns2 Инициализирует блок 2, сбрасывает состояния и обрабатывает его после проверки завершения блока 1.")
        Globals.userStates.remove(chatId)
        Globals.userPadezh.remove(chatId)
        Globals.userBlocks[chatId] = 2
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val (block1Completed, _, _) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
        if (block1Completed) {
            Globals.userBlocks[chatId] = 2
            handleBlock2(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
        } else {
            val messageText = "Вы не завершили блок \"Существительные\" 1.\nПройдите его перед переходом к блоку \"Существительные 2\"."
            bot.answerCallbackQuery(
                callbackQueryId = callbackQueryId,
                text = messageText,
                showAlert = true
            )
        }
    }

    fun callPadezh(chatId: Long, bot: Bot, data: String, callbackId: String) {
        println("Nouns2 callPadezh Устанавливает выбранный падеж, сбрасывает состояние и запускает обработку блока 2.")
        val selectedPadezh = data.removePrefix("nouns2Padezh:")
        Globals.userPadezh[chatId] = selectedPadezh
        Globals.userStates[chatId] = 0
        Globals.userColumnOrder.remove(chatId)
        handleBlock2(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    fun handleBlock2(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("Nouns2 handleBlock2 Обрабатывает блок 2: отправляет меню выбора или генерирует сообщение из Excel.")
        if (Globals.userPadezh[chatId] == null) {
            sendPadezhSelection(chatId, bot, filePath)
            return
        }
        if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
            sendWordMessage(chatId, bot, filePath)
            return
        }
        val blockRanges = Config.PADEZH_RANGES[Globals.userPadezh[chatId]] ?: return
        if (Globals.userColumnOrder[chatId].isNullOrEmpty()) {
            Globals.userColumnOrder[chatId] = blockRanges.shuffled().toMutableList()
        }
        val currentState = Globals.userStates[chatId] ?: 0
        val shuffledColumns = Globals.userColumnOrder[chatId]!!
        if (currentState >= shuffledColumns.size) {
            addScoreForPadezh(chatId, Globals.userPadezh[chatId].toString(), filePath, block = 2)
            return
        }
        val range = shuffledColumns[currentState]
        val messageText = generateMessageFromRange(filePath, "Существительные 2", range, wordUz, wordRus)
        val isLastMessage = currentState == shuffledColumns.size - 1
        if (isLastMessage) {
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    replyMarkup = Keyboards.nextCaseButtonWithHintToggleNouns2(wordUz, wordRus)
                )
            }
            addScoreForPadezh(chatId, Globals.userPadezh[chatId].toString(), filePath, block = 2)
        } else {
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    replyMarkup = Keyboards.nextButton(wordUz, wordRus)
                )
            }
        }
    }

    private fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
        println("Nouns2 sendPadezhSelection Формирует и отправляет клавиатуру для выбора падежа с учётом данных пользователя.")

        Globals.userPadezh.remove(chatId)
        Globals.userStates.remove(chatId)
        val currentBlock = Globals.userBlocks[chatId] ?: 2
        val padezhColumns = getPadezhColumnsForBlock(currentBlock)
        if (padezhColumns == null) {
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка: невозможно загрузить данные блока."
            )
            return
        }
        val userScores = getUserScoresForBlock(chatId, filePath, padezhColumns)
        val buttons = generatePadezhSelectionButtons(currentBlock, padezhColumns, userScores).toMutableList()

        if (Globals.userBlockCompleted[chatId]?.second == true) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("Перейти к Существительные 3", "nouns3")))
        }
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Выберите падеж для изучения блока $currentBlock:",
                replyMarkup = InlineKeyboardMarkup.create(buttons)
            )
        }
    }

    private fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
        println("Nouns2 getPadezhColumnsForBlock Возвращает диапазоны колонок для выбранного блока по падежам из конфигурации.")
        val result = Config.COLUMN_RANGES[block]
        return result
    }

    private fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
        println("Nouns2 getUserScoresForBlock Считывает баллы пользователя по падежам из Excel-листа состояния пользователя.")
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

    private fun generatePadezhSelectionButtons(currentBlock: Int, PadezhColumns: Map<String, Int>, userScores: Map<String, Int>): List<List<InlineKeyboardButton>> {
        println("Nouns2 generatePadezhSelectionButtons Формирует кнопки выбора падежей с отображением баллов пользователя.")
        val buttons = PadezhColumns.keys.map { PadezhName ->
            val score = userScores[PadezhName] ?: 0
            InlineKeyboardButton.CallbackData("$PadezhName [$score]", "nouns2Padezh:$PadezhName")
        }.map { listOf(it) }.toMutableList()
        return buttons
    }

    private fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
        println("Nouns2 sendWordMessage Отправляет клавиатуру для выбора слова, обрабатывая ошибки отсутствия файла.")
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

    private fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
        println("Nouns2 createWordSelectionKeyboardFromExcel Создаёт клавиатуру выбора слов, считывая случайные строки из Excel.")
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

    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?, showHint: Boolean = false): String {
        println("Nouns2 generateMessageFromRange Генерирует сообщение из указанного диапазона Excel, заменяя плейсхолдеры слов.")
        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        val excelManager = ExcelManager(filePath)
        val rawText = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName) ?: run {
                throw IllegalArgumentException("Лист $sheetName не найден")
            }
            val cells = extractCellsFromRange(sheet, range, wordUz)
            val mainText = cells.firstOrNull()?.replace("#", "$wordRus \\- $wordUz") ?: ""
            val hintText = cells.drop(1).joinToString("\n")
            val combinedText = if (hintText.isNotBlank()) "$mainText\n\n$hintText" else mainText
            combinedText
        }
        val finalText = if (showHint) {
            rawText.replace("\$\$", "")
        } else {
            rawText.replace(Regex("""\$\$.*?\$\$""", RegexOption.DOT_MATCHES_ALL), "$wordRus \\\\- $wordUz")
        }
        return finalText
    }

    private fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("Nouns2 extractCellsFromRange Извлекает и обрабатывает ячейки из диапазона для формирования сообщения.")
        val (start, end) = getRangeIndices(range)
        val column = range[0] - 'A'
        val cells = (start..end).mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex)
            if (row == null) {
                null
            } else {
                val processed = processRowForRange(row, column, wordUz)
                processed
            }
        }
        return cells
    }

    private fun getRangeIndices(range: String): Pair<Int, Int> {
        println("Nouns2 getRangeIndices Определяет начальный и конечный индексы строки из строкового диапазона.")
        val parts = range.split("-")
        if (parts.size != 2) {
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        return start to end
    }

    private fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        println("Nouns2 processRowForRange Обрабатывает строку таблицы, извлекая содержимое ячейки по диапазону.")
        val cell = row.getCell(column)
        val result = cell?.let { processCellContent(it, wordUz) }
        return result
    }

    private fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("Nouns2 processCellContent Обрабатывает содержимое ячейки, учитывая форматирование и заданное слово.")
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

    private fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("Nouns2 processFormattedRuns")
        val result = buildString {
            val runs = richText.numFormattingRuns()
            for (i in 0 until runs) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < runs) richText.getIndexOfFormattingRun(i + 1) else text.length
                val substring = text.substring(start, end)
                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
                val adjustedSubstring = adjustWordUz(substring, wordUz)
                val runContent = when (colorHex) {
                    "#FF0000" -> {
                        "||$adjustedSubstring||"
                    }
                    "#0000FF" -> {
                        "\$\$$adjustedSubstring\$\$"
                    }

                    else -> {
                        adjustedSubstring
                    }
                }
                append(runContent)
            }
        }
        val finalResult = result.escapeMarkdownV2()
        return finalResult
    }

    private fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("Nouns2 processCellWithoutRuns Обрабатывает ячейку без форматирования, применяя цветовые стили и экранирование.")
        val font = getCellFont(cell)
        val color = font?.let { getFontColor(it) }
        val rawContent = adjustWordUz(text, wordUz)
        val markedContent = when (color) {
            "#FF0000" -> "||$rawContent||"
            "#0000FF" -> "\$\$${rawContent}\$\$"
            else -> rawContent
        }
        val escapedContent = markedContent.escapeMarkdownV2()
        return escapedContent
    }

    private fun getCellFont(cell: Cell): XSSFFont? {
        println("Nouns2 getCellFont")
        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
        if (workbook == null) {
            return null
        }
        val fontIndex = cell.cellStyle.fontIndexAsInt
        val font = workbook.getFontAt(fontIndex) as? XSSFFont
        if (font != null) {
        } else {
        }
        return font
    }

    private fun getFontColor(font: XSSFFont): String {
        println("Nouns2 getFontColor")
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

    private fun adjustWordUz(content: String, wordUz: String?): String {
        println("Nouns2 adjustWordUz")
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
                            lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> {
                                "s"
                            }

                            lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> {
                                "i"
                            }

                            else -> {
                                ""
                            }
                        }
                        append(replacement)
                        append(nextChar)
                        i++
                    }

                    char == '*' -> {
                        append(wordUz)
                    }

                    else -> append(char)
                }
                i++
            }
        }
        return result
    }

    private fun String.escapeMarkdownV2(): String {
        println("Nouns2 escapeMarkdownV2 Экранирует специальные символы Markdown V2 в строке для корректного отображения.")
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
            //.replace("#", "\\#")
            .replace("+", "\\+")
            .replace("-", "\\-")
            .replace("=", "\\=")
            //.replace("|", "\\|")
            .replace("{", "\\{")
            .replace("}", "\\}")
            //.replace("´", "\\´")
            .replace(".", "\\.")
            .replace("!", "\\!")
        return escaped
    }

    private fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, block: Int) {
        println("Nouns2 addScoreForPadezh Добавляет балл пользователю, обновляя соответствующую ячейку в Excel файле.")

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

    private fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("Nouns2 getRgbWithTint Возвращает RGB значение с учётом оттенка, корректируя базовый цвет.")
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
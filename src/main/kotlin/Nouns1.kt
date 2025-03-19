import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.extractWordsFromCallback
import com.github.kotlintelegrambot.dispatcher.handleBlock
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
import kotlin.collections.mapNotNull
import kotlin.experimental.and

object Nouns1 {
    fun callNouns1(chatId: Long, bot: Bot) {
        println("Nouns1 callNouns1: Установлен блок 1 для $chatId")
        handleBlock1(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    fun handleBlock1(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("Nouns1 handleBlock1 // Обработка блока 1")
        if (Globals.userPadezh[chatId] == null) {
            sendPadezhSelection(chatId, bot, filePath)
            return
        }
        if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
            sendWordMessage(chatId, bot, Config.TABLE_FILE)
        } else {
            sendStateMessage(chatId, bot, Config.TABLE_FILE, wordUz, wordRus)
        }
    }

    fun callPadezh(chatId: Long, bot: Bot, data: String, callbackId: String) {
        println("Nouns1 callPadezh 23. Обработка выбора падежа")
        val selectedPadezh = data.removePrefix("nouns1Padezh:")
        Globals.userPadezh[chatId] = selectedPadezh
        Globals.userStates[chatId] = 0
        Globals.userColumnOrder.remove(chatId)
        bot.answerCallbackQuery(
            callbackQueryId = callbackId,
            text = """Падеж: $selectedPadezh
Слово: ${Globals.userWordUz[chatId]} - ${Globals.userWordRus[chatId]}""".trimMargin(),
            showAlert = true
        )
        handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
        println("Nouns1 sendPadezhSelection // Формирование клавиатуры для выбора падежа")
        Globals.userPadezh.remove(chatId)
        Globals.userStates.remove(chatId)
        val currentBlock = Globals.userBlocks[chatId] ?: 1
        val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
        if (PadezhColumns == null) {
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка: невозможно загрузить данные блока.")
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

    fun generatePadezhSelectionButtons(
        currentBlock: Int,
        PadezhColumns: Map<String, Int>,
        userScores: Map<String, Int>
    ): List<List<InlineKeyboardButton>> {
        println("Nouns1 generatePadezhSelectionButtons // Формирование кнопок выбора падежей")

        val buttons = PadezhColumns.keys.map { PadezhName ->
            val score = userScores[PadezhName] ?: 0
            InlineKeyboardButton.CallbackData("$PadezhName [$score]", "nouns1Padezh:$PadezhName")
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

    fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
        println("Nouns1 getPadezhColumnsForBlock // Получение колонок текущего блока")
        val result = Config.COLUMN_RANGES[block]
        return result
    }

    fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
        println("Nouns1 getUserScoresForBlock // Чтение баллов пользователя для блока")

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
                    }
                    break
                }
            }
        }

        return scores
    }

    fun processStateAndSendMessage(
        chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
        selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int
    ) {
        println("Nouns1 processStateAndSendMessage // Обработка состояния и отправка сообщения")

        if (currentState >= rangesForPadezh.size) {
            addScoreForPadezh(chatId, selectedPadezh, filePath)
            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
            return
        }

        val range = rangesForPadezh[currentState]
        val currentBlock = Globals.userBlocks[chatId] ?: 1
        val listName = when (currentBlock) {
            1 -> "Существительные 1"
            2 -> "Существительные 2"
            3 -> "Существительные 3"
            else -> "Существительные 1"
        }

        val messageText = try {
            generateMessageFromRange(filePath, listName, range, wordUz, wordRus)
        } catch (e: Exception) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
            return
        }

        sendMessageOrNextStep(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState, messageText)
    }

    fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("Nouns1 sendStateMessage // Отправка сообщения по текущему состоянию и падежу")

        val (selectedPadezh, rangesForPadezh, currentState) = validateUserState(chatId, bot) ?: return

        processStateAndSendMessage(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState)
    }

    fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
        println("Nouns1 sendFinalButtons // Отправка финального меню действий")

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
                replyMarkup = Keyboards.finalButtons(wordUz, wordRus, Globals.userBlocks[chatId] ?: 1)
            )
        }
    }

    fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, sheetName: String = "Состояние пользователя") {
        println("Nouns1 addScoreForPadezh // Добавляет балл для падежа")

        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден.")
        }

        val PadezhColumnIndex = when (Padezh) {
            "Именительный" -> 1
            "Родительный" -> 2
            "Винительный" -> 3
            "Дательный" -> 4
            "Местный" -> 5
            "Исходный" -> 6
            else -> throw IllegalArgumentException("A5 Ошибка: Неизвестный падеж: $Padezh")
        }

        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("A3 Ошибка: Лист $sheetName не найден")

            for (rowIndex in 1..sheet.lastRowNum) {
                val row = sheet.getRow(rowIndex) ?: continue
                val idCell = row.getCell(0)
                val currentId = when (idCell?.cellType) {
                    CellType.NUMERIC -> idCell.numericCellValue.toLong()
                    CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                    else -> null
                }
                if (currentId == chatId) {
                    val PadezhCell = row.getCell(PadezhColumnIndex) ?: row.createCell(PadezhColumnIndex)
                    val currentScore = PadezhCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                    PadezhCell.setCellValue(currentScore + 1)
                    excelManager.safelySaveWorkbook(workbook)
                    return@useWorkbook
                }
            }
        }
    }

    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?): String {
        println("Nouns1 generateMessageFromRange // Генерация сообщения из диапазона Excel")

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

    fun sendMessageOrNextStep(
        chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
        selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int, messageText: String
    ) {
        println("Nouns1 sendMessageOrNextStep // Отправка сообщения или переход к следующему шагу")

        if (currentState == rangesForPadezh.size - 1) {
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                //parseMode = ParseMode.MARKDOWN_V2
            )
            }
            val currentBlock = Globals.userBlocks[chatId] ?: 1
            com.github.kotlintelegrambot.dispatcher.addScoreForPadezh(chatId, selectedPadezh, filePath, currentBlock)
            com.github.kotlintelegrambot.dispatcher.sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
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
    fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("Nouns1 extractCellsFromRange // Извлечение и обработка ячеек из диапазона")

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

    fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("Nouns1 processCellContent // Обработка содержимого ячейки")

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

    fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("Nouns1 processCellWithoutRuns // Обработка ячейки без форматированных участков")

        val font = getCellFont(cell)

        val isRed = font != null && getFontColor(font) == "#FF0000"

        val result = if (isRed) {
            "||${adjustWordUz(text, wordUz).escapeMarkdownV2()}||"
        } else {
            adjustWordUz(text, wordUz).escapeMarkdownV2()
        }
        return result
    }

    fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
        println("Nouns1 sendWordMessage // Отправка клавиатуры с выбором слов")

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
        println("Nouns1 createWordSelectionKeyboardFromExcel // Создание клавиатуры из Excel-файла")
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

    // Проверка состояния пользователя
    fun validateUserState(chatId: Long, bot: Bot): Triple<String, List<String>, Int>? {
        println("Nouns1 validateUserState // Проверка состояния пользователя")
        val selectedPadezh = Globals.userPadezh[chatId]
        if (selectedPadezh == null) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
            return null
        }
        val rangesForPadezh = Config.PADEZH_RANGES[selectedPadezh]
        if (rangesForPadezh == null) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: диапазоны для падежа не найдены.")
            return null
        }

        val currentState = Globals.userStates[chatId] ?: 0
        return Triple(selectedPadezh, rangesForPadezh, currentState)
    }

    // Экранирование Markdown V2
    fun String.escapeMarkdownV2(): String {
        println("Nouns1 escapeMarkdownV2 // Экранирование Markdown V2")
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
        println("Nouns1 adjustWordUz // Обработка узбекского слова в контексте строки")

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

    // Обработка форматированных участков текста
    fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("Nouns1 processFormattedRuns // Обработка форматированных участков текста")

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
        println("Nouns1 getCellFont // Получение шрифта ячейки")

        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
        if (workbook == null) {
            return null
        }
        val fontIndex = cell.cellStyle.fontIndexAsInt

        val font = workbook.getFontAt(fontIndex) as? XSSFFont
        return font
    }

    // Функция для извлечения цвета шрифта
    fun getFontColor(font: XSSFFont): String {
        println("Nouns1 getFontColor // Извлечение цвета шрифта")

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
        println("Nouns1 getRgbWithTint // Получение RGB цвета с учётом оттенка")

        val baseRgb = rgb
        if (baseRgb == null) {
            println("g2 ⚠️ Базовый RGB не найден.")
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
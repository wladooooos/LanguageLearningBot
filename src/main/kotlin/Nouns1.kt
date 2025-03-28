import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.addScoreForPadezh
import com.github.kotlintelegrambot.dispatcher.handleBlock
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

object Nouns1 {
    fun callNouns1(chatId: Long, bot: Bot) {
        println("!!! Nouns1 callNouns1")
        Globals.userStates.remove(chatId)
        Globals.userPadezh.remove(chatId)
        Globals.userBlocks[chatId] = 1
        handleBlock1(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    fun handleBlock1(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("!!! Nouns1 handleBlock1")
        if (Globals.userPadezh[chatId] == null) {
            sendPadezhSelection(chatId, bot, filePath)
            return
        }
        val actualWordUz = wordUz ?: "bola"
        val actualWordRus = wordRus ?: "ребенок"
        Globals.userWordUz[chatId] = actualWordUz
        Globals.userWordRus[chatId] = actualWordRus
        sendStateMessage(chatId, bot, filePath, actualWordUz, actualWordRus)
    }

    fun callPadezh(chatId: Long, bot: Bot, data: String, callbackId: String) {
        println("!!! Nouns1 callPadezh")
        val selectedPadezh = data.removePrefix("nouns1Padezh:")
        Globals.userPadezh[chatId] = selectedPadezh
        Globals.userStates[chatId] = 0
        Globals.userColumnOrder.remove(chatId)
        handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    private fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
        println("!!! Nouns1 sendPadezhSelection")
        Globals.userPadezh.remove(chatId)
        Globals.userStates.remove(chatId)
        initializeUserBlockStates(chatId, filePath)
        val currentBlock = 1
        val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
        if (PadezhColumns == null) {
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка: невозможно загрузить данные блока."
            )
            return
        }
        val userScores = getUserScoresForBlock(chatId, filePath, PadezhColumns)
        val buttons = generatePadezhSelectionButtons(currentBlock, PadezhColumns, userScores).toMutableList()
        if (Globals.userBlockCompleted[chatId]?.first == true) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("Перейти к Существительные 2", "nouns2")))
        }
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Выберите падеж для изучения блока $currentBlock:",
                parseMode = "MarkdownV2",
                replyMarkup = InlineKeyboardMarkup.create(buttons)
            )
        }
    }

    private fun generatePadezhSelectionButtons(currentBlock: Int, PadezhColumns: Map<String, Int>, userScores: Map<String, Int>): List<List<InlineKeyboardButton>> {
        println("!!! Nouns1 generatePadezhSelectionButtons")
        val buttons = PadezhColumns.keys.map { PadezhName ->
            val score = userScores[PadezhName] ?: 0
            InlineKeyboardButton.CallbackData("$PadezhName [$score]", "nouns1Padezh:$PadezhName")
        }.map { listOf(it) }.toMutableList()
        return buttons
    }

    private fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
        println("!!! Nouns1 getPadezhColumnsForBlock")
        val result = Config.COLUMN_RANGES[block]
        if (result == null) {
        } else {
        }
        return result
    }

    private fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
        println("!!! Nouns1 getUserScoresForBlock")
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

    private fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("!!! Nouns1 sendStateMessage")
        val state = validateUserState(chatId, bot)
        if (state == null) {
            return
        }
        val (selectedPadezh, rangesForPadezh, currentState) = state
        val currentRange = rangesForPadezh[currentState]
        updateGlobals(chatId, currentRange)
        val messageText = generateMessageText(filePath, "Существительные 1", currentRange, wordUz, wordRus)
            ?: run {
                bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
                return
            }
        if (isLastStep(currentState, rangesForPadezh)) {
            handleLastStep(chatId, bot, filePath, messageText, selectedPadezh, wordUz, wordRus)
        } else {
            handleNextStep(chatId, bot, messageText, wordUz, wordRus)
        }
    }

    private fun updateGlobals(chatId: Long, currentRange: String) {
        println("!!! Nouns1 updateGlobals")
        Globals.currentSheetName[chatId] = "Существительные 1"
        Globals.currentRange[chatId] = currentRange
    }

    private fun generateMessageText(filePath: String, sheetName: String, currentRange: String, wordUz: String?, wordRus: String?): String? {
        println("!!! Nouns1 generateMessageText")
        return try {
            generateMessageFromRange(filePath, sheetName, currentRange, wordUz, wordRus, showHint = false)
        } catch (e: Exception) {
            null
        }
    }

    private fun isLastStep(currentState: Int, rangesForPadezh: List<String>): Boolean {
        println("!!! Nouns1 isLastStep")
        return currentState == rangesForPadezh.size - 1
    }

    private fun handleLastStep(chatId: Long, bot: Bot, filePath: String, messageText: String, selectedPadezh: String, wordUz: String?, wordRus: String?) {
        println("!!! Nouns1 handleLastStep")
        val nextPadezh = getNextPadezh(selectedPadezh)
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                replyMarkup = Keyboards.nextCaseButtonWithHintToggleNouns1(
                    wordUz,
                    wordRus,
                    isHintVisible = false,
                    blockId = "nouns1",
                )
            )
        }
        addScoreForPadezh(chatId, selectedPadezh, filePath, Globals.userBlocks[chatId] ?: 1)
    }

    private fun handleNextStep(chatId: Long, bot: Bot, messageText: String, wordUz: String?, wordRus: String?) {
        println("!!! Nouns1 handleNextStep")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                replyMarkup = Keyboards.nextButtonWithHintToggle(wordUz, wordRus, isHintVisible = false, blockId = "nouns1")
            )
        }
    }

    private fun getNextPadezh(currentPadezh: String): String {
        println("!!! Nouns1 getNextPadezh")
        val padezhOrder = listOf("Именительный", "Родительный", "Винительный", "Дательный", "Местный", "Исходный")
        val currentIndex = padezhOrder.indexOf(currentPadezh)
        return if (currentIndex != -1 && currentIndex < padezhOrder.size - 1) {
            padezhOrder[currentIndex + 1]
        } else {
            currentPadezh
        }
    }

    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?, showHint: Boolean = false): String {
        println("!!! Nouns1 generateMessageFromRange")
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
            val mainText = cells.firstOrNull() ?: ""
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
        println("!!! Nouns1 extractCellsFromRange")
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
        println("!!! Nouns1 getRangeIndices")
        val parts = range.split("-")
        if (parts.size != 2) {
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        return start to end
    }

    private fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        println("!!! Nouns1 processRowForRange")
        val cell = row.getCell(column)
        val result = cell?.let { processCellContent(it, wordUz) }
        return result
    }

    private fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("!!! Nouns1 processCellContent")
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

    private fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("!!! Nouns1 processCellWithoutRuns")
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

    fun validateUserState(chatId: Long, bot: Bot): Triple<String, List<String>, Int>? {
        println("!!! Nouns1 validateUserState")
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

    private fun String.escapeMarkdownV2(): String {
        println("!!! Nouns1 String.escapeMarkdownV2")
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
            //.replace("|", "\\|")
            .replace("{", "\\{")
            .replace("}", "\\}")
            .replace(".", "\\.")
            .replace("!", "\\!")
        return escaped
    }

    private fun adjustWordUz(content: String, wordUz: String?): String {
        println("!!! Nouns1 adjustWordUz")
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

    private fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("!!! Nouns1 processFormattedRuns")
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

    private fun getCellFont(cell: Cell): XSSFFont? {
        println("!!! Nouns1 getCellFont")
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
        println("!!! Nouns1 getFontColor")
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

    private fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("!!! Nouns1 XSSFColor.getRgbWithTint")
        val baseRgb = rgb
        if (baseRgb == null) {
            return null
        }
        val tint = this.tint
        val result = if (tint != 0.0) {
            val tinted = baseRgb.map { (it and 0xFF.toByte()) * (1 + tint) }
                .map { it.coerceIn(0.0, 255.0).toInt().toByte() }
                .toByteArray()
            tinted
        } else {
            baseRgb
        }
        return result
    }

    fun callRandomWord(chatId: Long, bot: Bot) {
        println("!!! Nouns1 callRandomWord")
        val filePath = Config.TABLE_FILE
        val file = File(filePath)
        if (!file.exists()) {
            bot.sendMessage(ChatId.fromId(chatId), text = "Ошибка: файл не найден.")
            return
        }
        val excelManager = ExcelManager(filePath)
        var randomWordUz = ""
        var randomWordRus = ""
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Существительные")
                ?: throw Exception("Лист 'Существительные' не найден")
            val randomRowIndex = (1..sheet.lastRowNum).random()
            val row = sheet.getRow(randomRowIndex)
            randomWordUz = row.getCell(0)?.toString()?.trim() ?: ""
            randomWordRus = row.getCell(1)?.toString()?.trim() ?: ""
        }
        Globals.userWordUz[chatId] = randomWordUz
        Globals.userWordRus[chatId] = randomWordRus
        handleBlock(chatId, bot, filePath, randomWordUz, randomWordRus)
    }
}
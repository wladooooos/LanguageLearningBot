import Adjectives2.markAdjective2AsCompleted
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
import kotlin.collections.set
import kotlin.experimental.and

object Adjectives1 {

    fun callAdjective1(chatId: Long, bot: Bot, callbackQueryId: String) {
        println("Adjectives1 callAdjective1: Запускает блок прилагательных 1 и устанавливает состояние.")
        println("main initializeUserBlockStates состояние ${Globals.userBlockCompleted[chatId]}!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val blocksCompleted = Globals.userBlockCompleted[chatId]
        if (blocksCompleted == null || !(blocksCompleted.first && blocksCompleted.second && blocksCompleted.third)) {
            Config.sendIncompleteBlocksAlert(bot, callbackQueryId)
            return
        }
        println("main initializeUserBlockStates состояние ${Globals.userBlockCompleted[chatId]}!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        Globals.userBlocks[chatId] = 5
        Globals.userStates.remove(chatId)
        Globals.userPadezh.remove(chatId)
        Globals.userColumnOrder.remove(chatId)
        handleBlockAdjective1(chatId, bot)
    }

    fun handleBlockAdjective1(chatId: Long, bot: Bot, keepState: Boolean = false) {
        println("Adjectives1 handleBlockAdjective1: Управляет логикой прохождения блока прилагательных 1.")
        initializeReplacementsIfNeeded(chatId)
        val rangesForAdjectives = Config.ADJECTIVE_RANGES_1
        val currentState = Globals.userStates[chatId] ?: 0
        // не сбрасываем состояние, если keepState == true
        if (!keepState) {
            Globals.userStates[chatId] = currentState
        }
        if (currentState >= rangesForAdjectives.size) {
            //sendFinalButtonsForAdjectives(chatId, bot)
            Globals.userReplacements.remove(chatId)
            return
        }
        val currentRange = rangesForAdjectives[currentState]
        Globals.currentSheetName[chatId] = "Прилагательные 1"
        Globals.currentRange[chatId] = currentRange
        val messageText = prepareAdjectiveMessage(chatId, bot, currentRange) ?: return
        val isLastRange = currentState == rangesForAdjectives.size - 1
        sendAdjectiveBlockMessage(chatId, bot, messageText, isLastRange)
    }

    fun callChangeWordsAdjective1(chatId: Long, bot: Bot) {
        println("Adjectives1 callChangeWordsAdjective1: Смена набора слов без сброса состояния.")
        Globals.userReplacements.remove(chatId)
        Globals.sheetColumnPairs.remove(chatId)
        initializeSheetColumnPairsFromFile(chatId)
        handleBlockAdjective1(chatId, bot, keepState = true)
    }


    private fun initializeReplacementsIfNeeded(chatId: Long) {
        println("Adjectives1 initializeReplacementsIfNeeded: Инициализирует замены, если они отсутствуют у пользователя.")
        if (Globals.userReplacements[chatId].isNullOrEmpty()) {
            initializeSheetColumnPairsFromFile(chatId)
        }
    }

    private fun prepareAdjectiveMessage(chatId: Long, bot: Bot, currentRange: String): String? {
        println("Adjectives1 prepareAdjectiveMessage: Генерирует сообщение для текущего мини-блока с обработкой ошибок.")
        return try {
            generateAdjectiveMessage(Config.TABLE_FILE, "Прилагательные 1", currentRange, Globals.userReplacements[chatId]!!)
        } catch (e: Exception) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
            null
        }
    }

    private fun sendAdjectiveBlockMessage(chatId: Long, bot: Bot, messageText: String, isLastRange: Boolean) {
        println("Adjectives1 sendAdjectiveBlockMessage: Отправляет сообщение с кнопкой 'Далее' или финальное меню.")

        val replyMarkup = if (isLastRange) {
            markAdjective1AsCompleted(chatId, Config.TABLE_FILE)
            Keyboards.adjective1HintToggleKeyboard(isHintVisible = false, isLast = true)
        } else {
            Keyboards.adjective1HintToggleKeyboard(isHintVisible = false, isLast = false)
        }

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessageWithoutMarkdown(
                chatId = chatId,
                text = messageText,
                replyMarkup = replyMarkup
            )
        }
    }



    fun generateReplacements(chatId: Long) {
        println("Adjectives1 generateReplacements: Формирует словарь замен на основе выбранных пар.")
        val userPairs = Globals.sheetColumnPairs[chatId] ?: run {
            return
        }
        val keysList = userPairs.keys.toList()
        if (keysList.size < 9) {
            println("⚠️ Предупреждение: у пользователя $chatId недостаточно пар для замены.")
        }
        val replacements = mutableMapOf<Int, String>()
        keysList.take(9).forEachIndexed { index, key ->
            replacements[index + 1] = key
        }
        Globals.userReplacements[chatId] = replacements
    }

    fun generateAdjectiveMessage(
        filePath: String,
        sheetName: String,
        range: String,
        replacements: Map<Int, String>,
        showHint: Boolean = false
    ): String {
        println("Adjectives1 generateAdjectiveMessage: Формирует текст сообщения для блока прилагательных с заменой цифр.")

        val sheetPairs = Globals.sheetColumnPairs.values.firstOrNull() ?: emptyMap()  // альтернатива, если chatId неизвестен
        val rawText = generateMessageFromRange(
            filePath = filePath,
            sheetName = sheetName,
            range = range,
            wordUz = null,
            wordRus = null,
            showHint = showHint,
            replacements = replacements,
            sheetColumnPairs = sheetPairs
        )

        val processedText = rawText.replace(Regex("[1-9]")) { match ->
            val digit = match.value.toInt()
            replacements[digit] ?: match.value
        }

        return processedText
    }



    // Отправка финального меню
//    fun sendFinalButtonsForAdjectives(chatId: Long, bot: Bot) {
//        println("Adjectives1 sendFinalButtonsForAdjectives: Отправляет финальное меню для завершения блока прилагательных.")
//        val currentBlock = Globals.userBlocks[chatId] ?: 5  // По умолчанию 5 (прилагательные 1)
////        val changeWordsCallback = if (currentBlock == 5) "change_words_adjective1" else "change_words_adjective2"
////
////        val navigationButton = if (currentBlock == 5) {
////            InlineKeyboardButton.CallbackData("Следующий блок", "block:adjective2")
////        } else {
////            InlineKeyboardButton.CallbackData("Предыдущий блок", "block:adjective1")
////        }
//        GlobalScope.launch {
//            TelegramMessageService.updateOrSendMessageWithoutMarkdown(
//            chatId = chatId,
//            text = "Вы завершили все этапы работы с этим блоком прилагательных. Что будем делать дальше?",
//            replyMarkup = Keyboards.finalAdjectiveButtons(Globals.userBlocks[chatId] ?: 5)
//        )
//        }
//    }

    fun generateMessageFromRange(
        filePath: String,
        sheetName: String,
        range: String,
        wordUz: String?,
        wordRus: String?,
        showHint: Boolean = false,
        replacements: Map<Int, String>,
        sheetColumnPairs: Map<String, String>
    ): String {
        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }

        val excelManager = ExcelManager(filePath)
        val result = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("Лист $sheetName не найден")
            val cells = extractCellsFromRange(sheet, range, wordUz)
            val firstCell = cells.firstOrNull() ?: ""
            val messageBody = cells.drop(1).joinToString("\n")
            listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
        }

        val finalText = if (showHint) {
            result.replace("\$\$", "")
        } else {
            val hintBlock = buildReplacementPairsBlock(replacements, sheetColumnPairs)
            result.replace(Regex("""\$\$.*?\$\$""", RegexOption.DOT_MATCHES_ALL), hintBlock)
        }

        return finalText
    }


    private fun buildReplacementPairsBlock(
        replacements: Map<Int, String>,
        sheetColumnPairs: Map<String, String>
    ): String {
        val pairs = replacements.entries.mapNotNull { (_, uzWord) ->
            val rusWord = sheetColumnPairs[uzWord]
            if (rusWord != null) "$rusWord – $uzWord" else null
        }
        return if (pairs.isNotEmpty()) pairs.joinToString("\n") else "*"
    }



//    fun String.escapeMarkdownV2(): String {
//        println("Adjectives1 String.escapeMarkdownV2: Экранирует спецсимволы Markdown V2 в строке.")
//        val escaped = this.replace("\\", "\\\\")
//            .replace("_", "\\_")
//            .replace("*", "\\*")
//            .replace("[", "\\[")
//            .replace("]", "\\]")
//            .replace("(", "\\(")
//            .replace(")", "\\)")
//            .replace("~", "\\~")
//            .replace("`", "\\`")
//            .replace(">", "\\>")
//            .replace("#", "\\#")
//            .replace("+", "\\+")
//            .replace("-", "\\-")
//            .replace("=", "\\=")
//            .replace("|", "\\|")
//            .replace("{", "\\{")
//            .replace("}", "\\}")
//            .replace(".", "\\.")
//            .replace("!", "\\!")
//        return escaped
//    }

    // Обработка узбекского слова в контексте строки
    fun adjustWordUz(content: String, wordUz: String?): String {
        println("Adjectives1 adjustWordUz: Обрабатывает узбекское слово, заменяя символы '+' и '*' по правилам.")
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
        println("Adjectives1 getRangeIndices: Определяет начальные и конечные индексы для диапазона строки.")
        val parts = range.split("-")
        if (parts.size != 2) {
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        return start to end
    }

    fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        println("Adjectives1 processRowForRange: Извлекает содержимое строки Excel и обрабатывает его.")
        val cell = row.getCell(column)
        return cell?.let { processCellContent(it, wordUz) }
    }

    fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("Adjectives1 extractCellsFromRange: Извлекает ячейки из заданного диапазона Excel.")
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

    // Обработка содержимого ячейки
    fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("Adjectives1 processCellContent: Обрабатывает содержимое ячейки с учетом форматирования.")
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
        println("Adjectives1 processCellWithoutRuns: Обрабатывает ячейку без форматированных участков, применяя стили.")
        val font = getCellFont(cell)
        val color = font?.let { getFontColor(it) }
        val rawContent = adjustWordUz(text, wordUz)
        val markedContent = when (color) {
            "#FF0000" -> "||$rawContent||"
            "#0000FF" -> "\$\$$rawContent\$\$"
            else -> rawContent
        }
        return markedContent//.escapeMarkdownV2()
    }

    // Обработка форматированных участков текста
    fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("Adjectives1 processFormattedRuns: Обрабатывает форматированные участки ячейки и выполняет замены.")
        val result = buildString {
            for (i in 0 until richText.numFormattingRuns()) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < richText.numFormattingRuns()) richText.getIndexOfFormattingRun(i + 1) else text.length
                val substring = text.substring(start, end)
                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
                val adjustedSubstring = adjustWordUz(substring, wordUz)
                when (colorHex) {
                    "#FF0000" -> append("||${adjustedSubstring/*.escapeMarkdownV2()*/}||")
                    "#0000FF" -> append("$$${adjustedSubstring/*.escapeMarkdownV2()*/}$$")
                    else -> append(adjustedSubstring/*.escapeMarkdownV2()*/)
                }
            }
        }
        return result
    }

    // Получение шрифта ячейки
    fun getCellFont(cell: Cell): XSSFFont? {
        println("Adjectives1 getCellFont: Получает шрифт ячейки из Excel-файла.")
        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
        if (workbook == null) {
            return null
        }
        val fontIndex = cell.cellStyle.fontIndexAsInt
        val font = workbook.getFontAt(fontIndex) as? XSSFFont
        return font
    }

    fun getFontColor(font: XSSFFont): String {
        println("Adjectives1 getFontColor: Извлекает HEX-код цвета шрифта ячейки.")
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

    fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("Adjectives1 XSSFColor.getRgbWithTint: Возвращает RGB цвета с учетом оттенка.")
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

    fun sendReplacementsMessage(chatId: Long, bot: Bot) {
        println("Adjectives1 sendReplacementsMessage: Отправляет сообщение с 9 парами слов для замены.")
        val userPairs = Globals.sheetColumnPairs[chatId]
        if (userPairs.isNullOrEmpty()) {
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessageWithoutMarkdown(
                    chatId = chatId,
                    text = "Ошибка: Данные для отправки не найдены."
                )
            }
            return
        }
        val messageText = userPairs.entries.joinToString("\n") { (key, value) ->
            "$key - $value"
        }
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessageWithoutMarkdown(
                chatId = chatId,
                text = messageText
            )
        }
    }

    fun initializeSheetColumnPairsFromFile(chatId: Long) {
        println("Adjectives1 initializeSheetColumnPairsFromFile: Инициализирует глобальные пары слов из Excel.")
        Globals.sheetColumnPairs[chatId] = mutableMapOf()
        val userPairs = collectSheetColumnPairs()
        if (userPairs.size == 9) {
            Globals.sheetColumnPairs[chatId] = userPairs
        }
        generateReplacements(chatId)
    }

    private fun collectSheetColumnPairs(): MutableMap<String, String> {
        println("Adjectives1 collectSheetColumnPairs: Собирает пары слов из листов Excel в один мап.")
        val userPairs = mutableMapOf<String, String>()
        val file = File(Config.TABLE_FILE)
        if (!file.exists()) {
            return userPairs
        }
        val excelManager = ExcelManager(Config.TABLE_FILE)
        val sheetNames = listOf("Существительные", "Глаголы", "Прилагательные")
        excelManager.useWorkbook { workbook ->
            for (sheetName in sheetNames) {
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    println("⚠️ Лист $sheetName не найден!")
                    continue
                }
                val selectedPairs = processSheetForPairs(sheet, sheetName)
                for ((key, value) in selectedPairs) {
                    userPairs[key] = value
                }
            }
        }
        return userPairs
    }


    private fun extractCandidatePairs(sheet: org.apache.poi.ss.usermodel.Sheet): List<Pair<String, String>> {
        println("Adjectives1 extractCandidatePairs: Извлекает все непустые пары из Excel-листа.")
        val candidates = mutableListOf<Pair<String, String>>()
        for (i in 0..sheet.lastRowNum) {
            val row = sheet.getRow(i) ?: continue
            val key = row.getCell(0)?.toString()?.trim() ?: ""
            val value = row.getCell(1)?.toString()?.trim() ?: ""
            if (key.isNotEmpty() && value.isNotEmpty()) {
                candidates.add(key to value)
            }
        }
        return candidates
    }

    private fun processSheetForPairs(sheet: org.apache.poi.ss.usermodel.Sheet, sheetName: String): List<Pair<String, String>> {
        println("Adjectives1 processSheetForPairs: Выбирает 3 случайные пары из кандидатов листа Excel.")
        val candidates = extractCandidatePairs(sheet)
        if (candidates.size < 3) {
            println("⚠️ Недостаточно данных в листе $sheetName для выбора 3 пар.")
            return emptyList()
        }
        return candidates.shuffled().take(3)
    }

    fun markAdjective1AsCompleted(chatId: Long, filePath: String) {
        println("🔹 markAdjective1AsCompleted: Помечает блок 'Прилагательные 1' как завершённый в листе 'Состояние пользователя'")
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя")
                ?: throw IllegalArgumentException("Лист 'Состояние пользователя' не найден")

            val userRow = sheet.find { row ->
                val idCell = row.getCell(0)
                val id = when (idCell?.cellType) {
                    CellType.NUMERIC -> idCell.numericCellValue.toLong()
                    CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                    else -> null
                }
                id == chatId
            }

            if (userRow != null) {
                val cell = userRow.getCell(14) ?: userRow.createCell(14) // O = 15-я колонка = индекс 14
                cell.setCellValue(1.0)
                excelManager.safelySaveWorkbook(workbook)
                println("✅ Прогресс по прилагательным 1 записан в колонку O.")
            } else {
                println("⚠️ Пользователь $chatId не найден на листе 'Состояние пользователя'")
            }
        }
    }

}
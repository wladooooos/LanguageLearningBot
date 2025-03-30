import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.initializeUserBlockStates
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.keyboards.Keyboards
import com.google.gson.Gson
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import kotlin.collections.joinToString
import kotlin.collections.map
import kotlin.collections.mapNotNull
import kotlin.collections.set
import kotlin.experimental.and

object Nouns3 {

    fun callNouns3(chatId: Long, bot: Bot) {
        println("nouns3 callNouns3 Инициализирует блок 3 после завершения блоков 1 и 2")
        Globals.userColumnOrder.remove(chatId)
        Globals.userStates.remove(chatId)
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val (block1Completed, block2Completed, _) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
        if (block1Completed && block2Completed) {
            Globals.userBlocks[chatId] = 3
            handleBlock3(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
        } else {
            val notCompletedBlocks = getIncompleteBlocks(block1Completed, block2Completed)
            sendIncompleteBlocksMessage(chatId, bot, notCompletedBlocks)
        }
    }


    private fun getIncompleteBlocks(block1Completed: Boolean, block2Completed: Boolean): List<String> {
        println("nouns3 getIncompleteBlocks Возвращает список незавершённых блоков")
        val notCompletedBlocks = mutableListOf<String>()
        if (!block1Completed) notCompletedBlocks.add("Блок 1")
        if (!block2Completed) notCompletedBlocks.add("Блок 2")
        return notCompletedBlocks
    }

    private fun sendIncompleteBlocksMessage(chatId: Long, bot: Bot, notCompletedBlocks: List<String>) {
        println("nouns3 sendIncompleteBlocksMessage Сообщает о незавершённых блоках")
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

    fun handleBlock3(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("nouns3 handleBlock3 Запускает логику блока 3 с проверкой диапазонов")
        val shuffledRanges = initializeBlock3Ranges(chatId)
        val currentState = Globals.userStates[chatId] ?: 0
        val completedRanges = getCompletedRanges(chatId, filePath)
        val currentRange = getCurrentRange(shuffledRanges, completedRanges.toList(), currentState)
        updateGlobalsForBlock3(chatId, currentRange)
        val messageText = generateBlock3Message(chatId, bot, filePath, currentRange, wordUz!!, wordRus!!)
            ?: return
        val isLastRange = currentState == 29 // Если 30 сообщений всего
        val replyMarkup = determineReplyMarkup(currentState, wordUz, wordRus)
        sendBlock3Message(chatId, messageText, replyMarkup)

        if (Globals.userBlockCompleted[chatId]?.third != true) {
            saveUserProgressBlok3(chatId, filePath, currentRange)
        }

//        updateProgressIfNeeded(chatId, filePath, currentState)
//        if (isLastRange) {
//            updateUserProgressForMiniBlocks(chatId, filePath, shuffledRanges.indices.toList())
//        }
    }

    private fun initializeBlock3Ranges(chatId: Long): MutableList<String> {
        println("nouns3 initializeBlock3Ranges Перемешивает диапазоны мини-блоков для блока 3")
        val allRanges = Config.ALL_RANGES_BLOCK_3
        if (Globals.userColumnOrder[chatId].isNullOrEmpty()) {
            Globals.userColumnOrder[chatId] = allRanges.shuffled().toMutableList()
        }
        var userColumnIndex = Globals.userColumnOrder[chatId]!!
        println("nouns3 initializeBlock3Ranges порядок вывода сообщений: $userColumnIndex")
        return userColumnIndex
    }

    private fun getCurrentRange(shuffledRanges: List<String>, completedRanges: List<String>, currentState: Int): String {
        println("nouns3 getCurrentRange Выбирает первый не пройденный диапазон")
        for (range in shuffledRanges) {
            if (!completedRanges.contains(range)) {
                return range
            }
        }
        return shuffledRanges[currentState % shuffledRanges.size]
    }

    private fun updateGlobalsForBlock3(chatId: Long, currentRange: String) {
        println("nouns3 updateGlobalsForBlock3 Обновляет глобальные переменные для блока 3")
        Globals.currentSheetName[chatId] = "Существительные 3"
        Globals.currentRange[chatId] = currentRange
    }

    private fun generateBlock3Message(chatId: Long, bot: Bot, filePath: String, currentRange: String, wordUz: String, wordRus: String): String? {
        println("nouns3 generateBlock3Message Генерирует сообщение для блока 3 из Excel")
        return try {
            generateMessageFromRange(filePath, "Существительные 3", currentRange, wordUz, wordRus, showHint = false)
        } catch (e: Exception) {
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
            null
        }
    }

    private fun determineReplyMarkup(currentState: Int, wordUz: String, wordRus: String): InlineKeyboardMarkup {
        return if (currentState == 29) {
            println("nouns3 determineReplyMarkup Формирует клавиатуру в зависимости от состояния")
            InlineKeyboardMarkup.create(
                listOf(
                    listOf(InlineKeyboardButton.CallbackData("Сменить слово", "nouns3Change_word_random")),
                    listOf(InlineKeyboardButton.CallbackData("Перейти к тесту", "block:test"))
                )
            )
        } else {
            Keyboards.buttonNextChengeWord(wordUz, wordRus)
        }
    }

    private fun sendBlock3Message(chatId: Long, messageText: String, replyMarkup: InlineKeyboardMarkup) {
        println("nouns3 sendBlock3Message Отправляет сообщение с клавиатурой в Telegram")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                replyMarkup = replyMarkup
            )
        }
    }

    private fun updateProgressIfNeeded(chatId: Long, filePath: String, currentState: Int) {
        println("nouns3 updateProgressIfNeeded Обновляет промежуточный прогресс мини-блоков")
        if ((currentState + 1) % 5 == 0) {
            println("📌 После ${currentState + 1} сообщения обновляем промежуточный прогресс")
            updateUserProgressForMiniBlocks(chatId, filePath, (0..currentState).toList())
        }
    }

    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?, showHint: Boolean = false): String {
        println("nouns3 generateMessageFromRange Генерирует сообщение из диапазона Excel")
        val excelManager = ExcelManager(filePath)
        val rawText = excelManager.useWorkbook { workbook ->
            val sheet = getSheetOrThrow(workbook, sheetName)
            val cells = extractCellsFromRange(sheet, range, wordUz)
            combineCells(cells, wordRus, wordUz)
        }
        return processRawText(rawText, showHint, wordRus, wordUz)
    }

    private fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("Nouns3 extractCellsFromRange Извлекает и обрабатывает ячейки из диапазона для формирования сообщения.")
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
    private fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        println("Nouns3 processRowForRange Обрабатывает строку таблицы, извлекая содержимое ячейки по диапазону.")
        val cell = row.getCell(column)
        val result = cell?.let { processCellContent(it, wordUz) }
        return result
    }
    private fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("Nouns3 processCellContent Обрабатывает содержимое ячейки, учитывая форматирование и заданное слово.")
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
        println("Nouns3 processFormattedRuns")
        val result = buildString {
            val runs = richText.numFormattingRuns()
            for (i in 0 until runs) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < runs) richText.getIndexOfFormattingRun(i + 1) else text.length
                val substring = text.substring(start, end)
                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
                //val adjustedSubstring = adjustWordUz(substring, wordUz)
                val runContent = when (colorHex) {
                    "#FF0000" -> {
                        "||$substring||"
                    }
                    "#0000FF" -> {
                        "\$\$$substring\$\$"
                    }

                    else -> {
                        substring
                    }
                }
                append(runContent)
            }
        }
        val finalAdjusted = adjustFinalText(result.toString(), wordUz)
        return finalAdjusted.escapeMarkdownV2()
    }
//    private fun adjustWordUz(content: String, wordUz: String?): String {
//        println("Nouns3 adjustWordUz")
//        fun Char.isVowel() = this.lowercaseChar() in "aeiouаеёиоуыэюя"
//        val result = buildString {
//            var i = 0
//            while (i < content.length) {
//                val char = content[i]
//                when {
//                    char == '+' && i + 1 < content.length -> {
//                        val nextChar = content[i + 1]
//                        val lastChar = wordUz?.lastOrNull()
//                        val replacement = when {
//                            lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> {
//                                "s"
//                            }
//
//                            lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> {
//                                "i"
//                            }
//
//                            else -> {
//                                ""
//                            }
//                        }
//                        append(replacement)
//                        append(nextChar)
//                        i++
//                    }
//
//                    char == '*' -> {
//                        append(wordUz)
//                    }
//
//                    else -> append(char)
//                }
//                i++
//            }
//        }
//        return result
//    }
private fun adjustFinalText(text: String, wordUz: String?): String {
    fun Char.isVowel() = this.lowercaseChar() in "aeiouаеёиоуыэюя"
    val sb = StringBuilder()
    var i = 0
    while (i < text.length) {
        when {
            // Замена символа '*' на wordUz
            text[i] == '*' -> {
                sb.append(wordUz)
                i++
            }
            // Обработка шаблона "+<символ>", при этом пропускаем маркеры "||"
            text[i] == '+' && i + 1 < text.length -> {
                var j = i + 1
                // Пропускаем все последовательности "||"
                while (j < text.length - 1 && text[j] == '|' && text[j + 1] == '|') {
                    j += 2
                }
                if (j < text.length) {
                    val nextChar = text[j]
                    val lastChar = wordUz?.lastOrNull()
                    val replacement = when {
                        lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> "s"
                        lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> "i"
                        else -> ""
                    }
                    sb.append(replacement)
                    sb.append(nextChar)
                    i = j + 1
                } else {
                    // Если после '+' нет значащего символа, просто добавляем '+'
                    sb.append('+')
                    i++
                }
            }
            // В остальных случаях копируем символ как есть
            else -> {
                sb.append(text[i])
                i++
            }
        }
    }
    return sb.toString()
}


    private fun getFontColor(font: XSSFFont): String {
        println("Nouns3 getFontColor")
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
    private fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("Nouns3 processCellWithoutRuns Обрабатывает ячейку без форматирования, применяя цветовые стили и экранирование.")
        val font = getCellFont(cell)
        val color = font?.let { getFontColor(it) }
        val rawContent = adjustFinalText(text, wordUz)
        val markedContent = when (color) {
            "#FF0000" -> "||$rawContent||"
            "#0000FF" -> "\$\$${rawContent}\$\$"
            else -> rawContent
        }
        val escapedContent = markedContent.escapeMarkdownV2()
        return escapedContent
    }
    private fun getCellFont(cell: Cell): XSSFFont? {
        println("Nouns3 getCellFont")
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
    private fun getRangeIndices(range: String): Pair<Int, Int> {
        println("Nouns3 getRangeIndices Определяет начальный и конечный индексы строки из строкового диапазона.")
        val parts = range.split("-")
        if (parts.size != 2) {
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        return start to end
    }

//    private fun validateFileExists(filePath: String): File {
//        println("nouns3 validateFileExists Проверяет наличие файла по указанному пути")
//        val file = File(filePath)
//        if (!file.exists()) {
//            throw IllegalArgumentException("Файл $filePath не найден")
//        }
//        return file
//    }

    private fun getSheetOrThrow(workbook: org.apache.poi.ss.usermodel.Workbook, sheetName: String): org.apache.poi.ss.usermodel.Sheet {
        println("nouns3 getSheetOrThrow Получает лист из Excel или выбрасывает исключение")
        return workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден")
    }

    private fun combineCells(cells: List<String>, wordRus: String?, wordUz: String?): String {
        println("nouns3 combineCells Объединяет ячейки в единое сообщение")
        val mainText = cells.firstOrNull()?.replace("#", "$wordRus \\- $wordUz") ?: ""
        val hintText = cells.drop(1).joinToString("\n")
        return if (hintText.isNotBlank()) "$mainText\n\n$hintText" else mainText
    }

    private fun processRawText(rawText: String, showHint: Boolean, wordRus: String?, wordUz: String?): String {
        println("nouns3 processRawText Обрабатывает текст, скрывая или показывая подсказки")
        return if (showHint) {
            rawText.replace("\$\$", "")
        } else {
            rawText.replace(Regex("""\$\$.*?\$\$""", RegexOption.DOT_MATCHES_ALL), "$wordRus \\\\- $wordUz")
        }
    }

//    fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
//        println("\n sendWordMessage // Отправка клавиатуры с выбором слов")
//        println("K1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath")
//        if (!isFileAvailable(filePath)) {
//            sendErrorMessage(chatId, bot, "Ошибка: файл с данными не найден.")
//            println("K3 Сообщение об ошибке отправлено пользователю $chatId")
//            return
//        }
//        val inlineKeyboard = try {
//            println("K4 Генерация клавиатуры из файла $filePath")
//            createWordSelectionKeyboardFromExcel(filePath, "Существительные")
//        } catch (e: Exception) {
//            println("K5 Ошибка при обработке данных: ${e.message}")
//            sendErrorMessage(chatId, bot, "Ошибка при обработке данных: ${e.message}")
//            println("K6 Сообщение об ошибке отправлено пользователю $chatId")
//            return
//        }
//        println("K7 Клавиатура успешно создана: $inlineKeyboard")
//        sendWordKeyboard(chatId, bot, inlineKeyboard)
//        println("K8 Сообщение с клавиатурой отправлено пользователю $chatId")
//    }
//
//    private fun isFileAvailable(filePath: String): Boolean {
//        return File(filePath).exists()
//    }
//
//    private fun sendErrorMessage(chatId: Long, bot: Bot, message: String) {
//        GlobalScope.launch {
//            TelegramMessageService.updateOrSendMessage(
//                chatId = chatId,
//                text = message
//            )
//        }
//    }
//
//    private fun sendWordKeyboard(chatId: Long, bot: Bot, inlineKeyboard: InlineKeyboardMarkup) {
//        GlobalScope.launch {
//            TelegramMessageService.updateOrSendMessage(
//                chatId = chatId,
//                text = "Выберите слово из списка:",
//                replyMarkup = inlineKeyboard
//            )
//        }
//    }
//
//
//    // Создание клавиатуры из Excel-файла
//    fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
//        println("\n createWordSelectionKeyboardFromExcel // Создание клавиатуры из Excel-файла")
//        println("L1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName")
//
//        println("L2 Проверка существования файла $filePath")
//        val file = File(filePath)
//        if (!file.exists()) {
//            println("L3 Ошибка: файл $filePath не найден")
//            throw IllegalArgumentException("Файл $filePath не найден")
//        }
//
//        val excelManager = ExcelManager(filePath)
//        val buttons = excelManager.useWorkbook { workbook ->
//            val sheet = workbook.getSheet(sheetName)
//                ?: throw IllegalArgumentException("D5 Ошибка: Лист $sheetName не найден")
//            println("L6 Лист найден: $sheetName")
//            println("L7 Всего строк в листе: ${sheet.lastRowNum + 1}")
//            val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
//            println("L8 Случайные строки для кнопок: $randomRows")
//            println("L9 Начало обработки строк для создания кнопок")
//            val buttons = randomRows.mapNotNull { rowIndex ->
//                val row = sheet.getRow(rowIndex)
//                if (row == null) {
//                    println("L10 Строка $rowIndex отсутствует, пропускаем")
//                    return@mapNotNull null
//                }
//
//                val wordUz = row.getCell(0)?.toString()?.trim()
//                val wordRus = row.getCell(1)?.toString()?.trim()
//
//                if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
//                    println("L11 Некорректные данные в строке $rowIndex: wordUz=$wordUz, wordRus=$wordRus. Пропускаем")
//                    return@mapNotNull null
//                }
//
//                println("L12 Обработана строка $rowIndex: wordUz=$wordUz, wordRus=$wordRus")
//                InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
//            }.chunked(2)
//            println("L13 Кнопки успешно созданы. Количество строк кнопок: ${buttons.size}")
//            buttons
//        }
//        println("L14 Файл Excel закрыт. Генерация клавиатуры завершена")
//
//        return InlineKeyboardMarkup.create(buttons)
//    }
//
//    // Отправка финального меню действий
//    fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
//        println("\nR sendFinalButtons // Отправка финального меню действий")
//        println("R1 Вход в функцию. Параметры: chatId=$chatId, wordUz=$wordUz, wordRus=$wordRus")
//
//        GlobalScope.launch {
//            TelegramMessageService.updateOrSendMessage(
//                chatId = chatId,
//                text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
//                replyMarkup = Keyboards.finalButtons(wordUz, wordRus, Globals.userBlocks[chatId] ?: 1)
//            )
//        }
//        println("R2 Финальное меню отправлено пользователю $chatId")
//    }

    // Генерация сообщения из диапазона Excel

//    fun generateMessageFromRange(
//        filePath: String,
//        sheetName: String,
//        range: String,
//        wordUz: String?,
//        wordRus: String?,
//        showHint: Boolean = false
//    ): String {
//        println("\nNouns3 generateMessageFromRange // Генерация сообщения из диапазона Excel")
//        val file = File(filePath)
//        if (!file.exists()) throw IllegalArgumentException("Файл $filePath не найден")
//        val excelManager = ExcelManager(filePath)
//        // Формируем исходный текст, который может содержать spoiler‑маркировку (||...||)
//        val rawText = excelManager.useWorkbook { workbook ->
//            val sheet = workbook.getSheet(sheetName)
//                ?: throw IllegalArgumentException("Лист $sheetName не найден")
//            val cells = extractCellsFromRange(sheet, range, wordUz)
//            val firstCell = cells.firstOrNull() ?: ""
//            val messageBody = cells.drop(1).joinToString("\n")
//            listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
//        }
//        return if (showHint) {
//            // Если подсказка показывается – удаляем все маркеры "||"
//            rawText.replace("||", "")
//        } else {
//            // Если подсказка скрыта – заменяем каждый блок вида "||...||" (с переводами строк) на символ "*"
//            rawText.replace(Regex("\\|\\|.*?\\|\\|", RegexOption.DOT_MATCHES_ALL), "*")
//        }
//    }


//    Старая:
//    fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?): String {
//        println("SSS generateMessageFromRange // Генерация сообщения из диапазона Excel")
//        println("S1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName, range=$range, wordUz=$wordUz, wordRus=$wordRus")
//
//        val file = File(filePath)
//        if (!file.exists()) {
//            println("S2 Ошибка: файл $filePath не найден.")
//            throw IllegalArgumentException("Файл $filePath не найден")
//        }
//
//        val excelManager = ExcelManager(filePath)
//        val result = excelManager.useWorkbook { workbook ->
//            val sheet = workbook.getSheet(sheetName)
//            if (sheet == null) {
//                println("S3 Ошибка: лист $sheetName не найден.")
//                throw IllegalArgumentException("Лист $sheetName не найден")
//            }
//
//            println("S4 Лист $sheetName найден. Извлекаем ячейки из диапазона $range")
//            val cells = extractCellsFromRange(sheet, range, wordUz)
//            println("S5 Извлеченные ячейки: $cells")
//
//            val firstCell = cells.firstOrNull() ?: ""
//            val messageBody = cells.drop(1).joinToString("\n")
//
//            val res = listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
//            println("S7 Генерация завершена. Результат: $res")
//            res
//        }
//        println("S6 Файл Excel закрыт.")
//        return result
//    }

    // Экранирование Markdown V2
    private fun String.escapeMarkdownV2(): String {
        println("nouns3 escapeMarkdownV2 Экранирует спецсимволы для Markdown V2")
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
            //.replace("+", "\\+")
            .replace("-", "\\-")
            .replace("=", "\\=")
            //.replace("|", "\\|")
            .replace("{", "\\{")
            .replace("}", "\\}")
            .replace(".", "\\.")
            .replace("!", "\\!")
        return escaped
    }

//    fun adjustWordUz(content: String, wordUz: String?): String {
//        println("nouns3 adjustWordUz Обрабатывает узбекское слово в строке")
//        fun Char.isVowel() = this.lowercaseChar() in "aeiouаеёиоуыэюя"
//        val result = buildString {
//            var i = 0
//            while (i < content.length) {
//                val char = content[i]
//                when {
//                    char == '+' && i + 1 < content.length -> {
//                        val nextChar = content[i + 1]
//                        val lastChar = wordUz?.lastOrNull()
//                        val replacement = when {
//                            lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> "s"
//                            lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> "i"
//                            else -> ""
//                        }
//                        append(replacement)
//                        append(nextChar)
//                        i++
//                    }
//                    char == '*' -> append(wordUz)
//                    else -> append(char)
//                }
//                i++
//            }
//        }
//        return result
//    }

//    fun getRangeIndices(range: String): Pair<Int, Int> {
//        println("nouns3 getRangeIndices Вычисляет индексы начала и конца диапазона")
//        val parts = range.split("-")
//        if (parts.size != 2) {
//            throw IllegalArgumentException("Неверный формат диапазона: $range")
//        }
//        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
//        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
//        return start to end
//    }
//
//    fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
//        println("nouns3 processRowForRange Обрабатывает строку, извлекая содержимое ячейки")
//        val cell = row.getCell(column)
//        return cell?.let { processCellContent(it, wordUz) }
//    }

//    fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
//        println("\n extractCellsFromRange // Извлечение и обработка ячеек из диапазона")
//        println("V1 Вход в функцию. Параметры: range=\"$range\", wordUz=\"$wordUz\"")
//
//        val (start, end) = getRangeIndices(range)
//        val column = range[0] - 'A'
//        println("V2 Диапазон строк: $start-$end, Колонка: $column")
//
//        return (start..end).mapNotNull { rowIndex ->
//            val row = sheet.getRow(rowIndex)
//            if (row == null) {
//                println("V3 ⚠️ Строка $rowIndex отсутствует, пропускаем.")
//                null
//            } else {
//                val processed = processRowForRange(row, column, wordUz)
//                println("V5 ✅ Обработанная ячейка в строке $rowIndex: \"$processed\"")
//                processed
//            }
//        }.also { cells ->
//            println("V6 Результат извлечения: $cells")
//        }
//    }

//    // Обработка содержимого ячейки
//    fun processCellContent(cell: Cell?, wordUz: String?): String {
//        println("nouns3 processCellContent Обрабатывает содержимое ячейки с форматированием")
//        if (cell == null) {
//            return ""
//        }
//        val richText = cell.richStringCellValue as XSSFRichTextString
//        val text = richText.string
//        val runs = richText.numFormattingRuns()
//        val result = if (runs == 0) {
//            processCellWithoutRuns(cell, text, wordUz)
//        } else {
//            processFormattedRuns(richText, text, wordUz)
//        }
//        return result
//    }

//    fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
//        println("nouns3 processCellWithoutRuns Обрабатывает ячейку без форматированных участков")
//        val font = getCellFont(cell)
//        val isRed = font != null && getFontColor(font) == "#FF0000"
//        val result = if (isRed) {
//            "||${adjustWordUz(text, wordUz).escapeMarkdownV2()}||"
//        } else {
//            adjustWordUz(text, wordUz).escapeMarkdownV2()
//        }
//        return result
//    }
//    fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
//        println("nouns3 processFormattedRuns Обрабатывает форматированные участки текста")
//        val result = buildString {
//            for (i in 0 until richText.numFormattingRuns()) {
//                val start = richText.getIndexOfFormattingRun(i)
//                val end = if (i + 1 < richText.numFormattingRuns()) richText.getIndexOfFormattingRun(i + 1) else text.length
//                val substring = text.substring(start, end)
//                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
//                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
//                val adjustedSubstring = adjustWordUz(substring, wordUz)
//                if (colorHex == "#FF0000") {
//                    append("||${adjustedSubstring.escapeMarkdownV2()}||")
//                } else {
//                    append(adjustedSubstring.escapeMarkdownV2())
//                }
//            }
//        }
//        return result
//    }

//    fun getCellFont(cell: Cell): XSSFFont? {
//        println("nouns3 getCellFont Получает шрифт ячейки из Excel")
//        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
//        if (workbook == null) {
//            return null
//        }
//        val fontIndex = cell.cellStyle.fontIndexAsInt
//        val font = workbook.getFontAt(fontIndex) as? XSSFFont
//        return font
//    }
//
//    fun getFontColor(font: XSSFFont): String {
//        println("nouns3 getFontColor Извлекает HEX цвет шрифта из ячейки")
//        val xssfColor = font.xssfColor
//        if (xssfColor == null) {
//            return "Цвет не определён"
//        }
//        val rgb = xssfColor.rgb
//        val result = if (rgb != null) {
//            val colorHex = rgb.joinToString(prefix = "#", separator = "") { "%02X".format(it) }
//            colorHex
//        } else {
//            "Цвет не определён"
//        }
//        return result
//    }

    private fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("nouns3 XSSFColor.getRgbWithTint Возвращает RGB цвет с учётом оттенка")
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

    private fun updateUserProgressForMiniBlocks(chatId: Long, filePath: String, completedMiniBlocks: List<Int>) {
        println("nouns3 updateUserProgressForMiniBlocks Обновляет прогресс мини-блоков пользователя")
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = getUserStateSheet(workbook)
            val userRow = findUserRow(sheet, chatId)
            if (userRow != null) {
                updateMiniBlockCells(userRow, completedMiniBlocks)
                excelManager.safelySaveWorkbook(workbook)
            } else {
            }
        }
    }

    private fun getUserStateSheet(workbook: org.apache.poi.ss.usermodel.Workbook): org.apache.poi.ss.usermodel.Sheet {
        println("nouns3 getUserStateSheet Получает лист состояния пользователя из Excel")
        return workbook.getSheet("Состояние пользователя")
            ?: throw IllegalArgumentException("Лист 'Состояние пользователя' не найден.")
    }

    private fun findUserRow(sheet: org.apache.poi.ss.usermodel.Sheet, chatId: Long): org.apache.poi.ss.usermodel.Row? {
        println("nouns3 findUserRow Ищет строку пользователя по chatId в Excel")
        for (row in sheet) {
            val idCell = row.getCell(0)
            val chatIdFromCell = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }
            if (chatIdFromCell == chatId) {
                return row
            }
        }
        return null
    }

    private fun updateMiniBlockCells(row: org.apache.poi.ss.usermodel.Row, completedMiniBlocks: List<Int>) {
        println("nouns3 updateMiniBlockCells Обновляет ячейки мини-блоков в строке пользователя")
        completedMiniBlocks.forEach { miniBlock ->
            val columnIndex = 11 + miniBlock // L = 11, M = 12, и т.д.
            val cell = row.getCell(columnIndex) ?: row.createCell(columnIndex)
            val currentValue = cell.numericCellValue.takeIf { it > 0 } ?: 0.0
            cell.setCellValue(currentValue + 1)
        }
    }

    private fun saveUserProgressBlok3(chatId: Long, filePath: String, range: String) {
        println("nouns3 saveUserProgressBlok3 Сохраняет прогресс блока 3 в Excel")
        val columnIndex = getColumnIndexForRange(range)
        if (columnIndex == null) {
            return
        }
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = getOrCreateUserStateSheet(workbook, "Состояние пользователя 3 блок")
            val userColumnIndex = findOrCreateUserColumnIndex(sheet, chatId)
            if (userColumnIndex == -1) {
                return@useWorkbook
            }
            updateProgressCell(sheet, userColumnIndex, columnIndex, chatId)
            excelManager.safelySaveWorkbook(workbook)
        }
    }

    private fun getColumnIndexForRange(range: String): Int? {
        println("nouns3 getColumnIndexForRange Определяет индекс столбца по диапазону")
        return Config.COLUMN_MAPPING[range]
    }

    private fun getOrCreateUserStateSheet(workbook: org.apache.poi.ss.usermodel.Workbook, sheetName: String): org.apache.poi.ss.usermodel.Sheet {
        println("nouns3 getOrCreateUserStateSheet Получает или создаёт лист состояния пользователя")
        return workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName)
    }

    private fun findOrCreateUserColumnIndex(sheet: org.apache.poi.ss.usermodel.Sheet, chatId: Long): Int {
        println("nouns3 findOrCreateUserColumnIndex Находит или создаёт ячейку для chatId в заголовке")
        val headerRow = sheet.getRow(0) ?: sheet.createRow(0)
        for (col in 0 until 31) { // Проверяем до 30 столбца
            val cell = headerRow.getCell(col) ?: headerRow.createCell(col)
            when (cell.cellType) {
                CellType.NUMERIC -> {
                    if (cell.numericCellValue.toLong() == chatId) {
                        return col
                    }
                }
                CellType.BLANK -> {
                    cell.setCellValue(chatId.toDouble())
                    return col
                }
                else -> continue
            }
        }
        return -1
    }

    private fun updateProgressCell(sheet: org.apache.poi.ss.usermodel.Sheet, userColumnIndex: Int, rowIndex: Int, chatId: Long) {
        println("nouns3 updateProgressCell Обновляет значение прогресса в ячейке Excel")
        val row = sheet.getRow(rowIndex) ?: sheet.createRow(rowIndex)
        val cell = row.getCell(userColumnIndex) ?: row.createCell(userColumnIndex)
        val currentScore = cell.numericCellValue.takeIf { it > 0 } ?: 0.0
        cell.setCellValue(currentScore + 1)
    }


    private fun getCompletedRanges(chatId: Long, filePath: String): Set<String> {
        println("nouns3 getCompletedRanges Извлекает завершённые диапазоны мини-блоков")
        val completedRanges = mutableSetOf<String>()
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = getUserStateSheet(workbook, "Состояние пользователя 3 блок")
            val headerRow = sheet.getRow(0) ?: return@useWorkbook
            val userColumnIndex = findUserColumnIndexInHeader(headerRow, chatId)
            if (userColumnIndex == null) {
                return@useWorkbook
            }
            val rangeMapping = Config.ALL_RANGES_BLOCK_3
            extractCompletedRanges(sheet, userColumnIndex, rangeMapping, completedRanges)
        }
        return completedRanges
    }


    private fun getUserStateSheet(workbook: org.apache.poi.ss.usermodel.Workbook, sheetName: String): org.apache.poi.ss.usermodel.Sheet {
        println("nouns3 getUserStateSheet Получает или создаёт лист с указанным именем из книги Excel")
        return workbook.getSheet(sheetName) ?: return workbook.createSheet(sheetName)
    }

    private fun findUserColumnIndexInHeader(headerRow: org.apache.poi.ss.usermodel.Row, chatId: Long): Int? {
        println("nouns3 findUserColumnIndexInHeader Находит столбец пользователя в заголовке листа")
        for (col in 0 until 31) {
            val cell = headerRow.getCell(col)
            if (cell != null && cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                return col
            }
        }
        return null
    }

    private fun extractCompletedRanges(sheet: org.apache.poi.ss.usermodel.Sheet, userColumnIndex: Int, rangeMapping: List<String>, completedRanges: MutableSet<String>) {
        println("nouns3 extractCompletedRanges Извлекает завершённые диапазоны из Excel-листа")
        for ((index, range) in rangeMapping.withIndex()) {
            val row = sheet.getRow(index + 1) ?: continue
            val cell = row.getCell(userColumnIndex)
            val value = cell?.numericCellValue ?: 0.0
            if (value > 0) {
                completedRanges.add(range)
            }
        }
    }


    fun callRandomWord(chatId: Long, bot: Bot) {
        println("nouns3 callRandomWord Выбирает случайное слово и обновляет сообщение")
        val filePath = Config.TABLE_FILE
        val (randomWordUz, randomWordRus) = pickRandomWordFromSheet(filePath, "Существительные")
        updateGlobalWords(chatId, randomWordUz, randomWordRus)
        handleBlock3ChangeWord(chatId, bot, filePath)
    }

//    private fun validateExcelFile(filePath: String, chatId: Long, bot: Bot): File? {
//        val file = File(filePath)
//        if (!file.exists()) {
//            println("callRandomWord: Файл не найден по пути $filePath")
//            bot.sendMessage(ChatId.fromId(chatId), text = "Ошибка: файл не найден.")
//            return null
//        }
//        println("callRandomWord: Файл найден по пути $filePath")
//        return file
//    }

    private fun pickRandomWordFromSheet(filePath: String, sheetName: String): Pair<String, String> {
        println("nouns3 pickRandomWordFromSheet Извлекает случайное слово с листа Excel")
        val excelManager = ExcelManager(filePath)
        var randomWordUz = ""
        var randomWordRus = ""
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName) ?: throw Exception("Лист '$sheetName' не найден")
            val randomRowIndex = (1..sheet.lastRowNum).random()
            val row = sheet.getRow(randomRowIndex)
            randomWordUz = row.getCell(0)?.toString()?.trim() ?: ""
            randomWordRus = row.getCell(1)?.toString()?.trim() ?: ""
        }
        return Pair(randomWordUz, randomWordRus)
    }

    private fun updateGlobalWords(chatId: Long, randomWordUz: String, randomWordRus: String) {
        println("nouns3 updateGlobalWords Обновляет глобальные переменные выбранными словами")
        Globals.userWordUz[chatId] = randomWordUz
        Globals.userWordRus[chatId] = randomWordRus
    }

    private fun handleBlock3ChangeWord(chatId: Long, bot: Bot, filePath: String) {
        println("nouns3 handleBlock3ChangeWord Обновляет сообщение блока 3 новым словом")
        if (!areWordsSelected(chatId, bot)) return
        val (currentSheetName, currentRange) = getCurrentSheetAndRange(chatId, bot) ?: return
        val newMessageText = generateNewMessageText(filePath, currentSheetName, currentRange,
            Globals.userWordUz[chatId], Globals.userWordRus[chatId], bot, chatId) ?: return
        val mainMessageId = getMainMessageId(chatId) ?: run {
            return
        }
        updateMessage(chatId, mainMessageId, newMessageText, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    private fun areWordsSelected(chatId: Long, bot: Bot): Boolean {
        println("nouns3 areWordsSelected Проверяет, что слова выбраны пользователем")
        if (Globals.userWordUz[chatId].isNullOrBlank() || Globals.userWordRus[chatId].isNullOrBlank()) {
            return false
        }
        return true
    }

    private fun getCurrentSheetAndRange(chatId: Long, bot: Bot): Pair<String, String>? {
        println("nouns3 getCurrentSheetAndRange Получает текущий лист и диапазон из Globals")
        val currentSheetName = Globals.currentSheetName[chatId] ?: "Существительные 3"
        val currentRange = Globals.currentRange[chatId]
        if (currentRange == null) {
            bot.sendMessage(ChatId.fromId(chatId), text = "Ошибка: текущий диапазон не найден.")
            return null
        }
        return Pair(currentSheetName, currentRange)
    }

    private fun generateNewMessageText(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?, bot: Bot, chatId: Long): String? {
        println("nouns3 generateNewMessageText Генерирует новый текст сообщения из Excel")
        return try {
            generateMessageFromRange(filePath, sheetName, range, wordUz, wordRus, showHint = false)
        } catch (e: Exception) {
            bot.sendMessage(ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
            null
        }
    }

    private fun getMainMessageId(chatId: Long): String? {
        println("nouns3 getMainMessageId Возвращает ID основного сообщения для chatId")
        return Globals.userMainMessageId[chatId]?.toString()
    }

    private fun updateMessage(chatId: Long, mainMessageId: String, newMessageText: String, wordUz: String?, wordRus: String?) {
        println("nouns3 updateMessage Редактирует сообщение с новым текстом и клавиатурой")
        val replacedMessageText = newMessageText.replace("#", "$wordRus - $wordUz")
        GlobalScope.launch {
            TelegramMessageService.editMessageTextViaHttpSuspend(
                chatId = chatId,
                messageId = mainMessageId.toLong(),
                text = replacedMessageText,
                replyMarkupJson = Gson().toJson(Keyboards.buttonNextChengeWord(wordUz, wordRus))
            )
        }
    }
}
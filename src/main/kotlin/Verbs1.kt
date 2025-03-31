import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.initializeUserBlockStates
import com.github.kotlintelegrambot.keyboards.Keyboards
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import kotlin.collections.joinToString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

object Verbs1 {
    fun callVerbs1(chatId: Long, bot: Bot, callbackQueryId: String) {
        println("callVerbs1: Обработка блока глаголов 1")
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val blocksCompleted = Globals.userAdjectiveCompleted[chatId]
        if (blocksCompleted == null || !(blocksCompleted.first) || !(blocksCompleted.second)) {
            Config.sendIncompleteBlocksAlert(bot, callbackQueryId)
            return
        }
        Globals.userVerb.remove(chatId)
        Globals.userBlocks[chatId] = 7
        Globals.userStates[chatId] = 0
        handleBlockVerbs1(chatId, bot)
    }

    fun handleBlockVerbs1(chatId: Long, bot: Bot, keepState: Boolean = false) {
        val setup = part1Setup(chatId, keepState)
        val showHint = setup[0] as Boolean
        val filePath = setup[1] as String
        val sheetName = setup[2] as String
        val currentColumnIndex = setup[3] as Int
        val initialMessage = part2BuildInitialMessage(chatId, filePath, sheetName, currentColumnIndex, showHint)
        val finalize = part3FinalizeMessage(chatId, initialMessage, currentColumnIndex, showHint)
        val finalMessageText = finalize[0] as String
        val isLastRange = finalize[1] as Boolean
        val replyMarkup = finalize[2] as com.github.kotlintelegrambot.entities.InlineKeyboardMarkup?
        part4SendAndComplete(chatId, finalMessageText, replyMarkup, isLastRange)
    }

    private fun part1Setup(chatId: Long, keepState: Boolean): List<Any> {
        val showHint = Globals.userVerbsHintVisibility[chatId] ?: false
        val filePath = "Глаголы.xlsx"
        val sheetName = "Глаголы 1"
        val currentColumnIndex = Globals.userStates[chatId] ?: 0
        if (!keepState) {
            Globals.userStates[chatId] = currentColumnIndex
        }
        return listOf(showHint, filePath, sheetName, currentColumnIndex)
    }

    private fun part2BuildInitialMessage(chatId: Long, filePath: String, sheetName: String, currentColumnIndex: Int, showHint: Boolean): String {
        val lines = read7LinesFromColumnWithBlueHint(filePath, sheetName, currentColumnIndex, showHint)
        val messageText = buildVerbsMessage(lines, chatId)
        return messageText
    }

    private fun part3FinalizeMessage(
        chatId: Long,
        messageText: String,
        currentColumnIndex: Int,
        showHint: Boolean
    ): List<Any> {
        val isLastRange = (currentColumnIndex == 15)
        println("isLastColumn = $isLastRange")
        val finalMessageText = replaceHashInMessage(chatId, messageText)
        println("Финальное сообщение после замены #: \n$finalMessageText")
        val replyMarkup = if (isLastRange) {
            Keyboards.verbs1HintToggleKeyboard(isHintVisible = showHint, isLast = true)
        } else {
            Keyboards.verbs1HintToggleKeyboard(isHintVisible = showHint, isLast = false)
        }
        return listOf(finalMessageText, isLastRange, replyMarkup)
    }

    private fun part4SendAndComplete(chatId: Long, messageText: String, replyMarkup: com.github.kotlintelegrambot.entities.InlineKeyboardMarkup?, isLastRange: Boolean) {
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessageWithoutMarkdown(
                chatId = chatId,
                text = messageText,
                replyMarkup = replyMarkup
            )
        }
        if (isLastRange) {
            markVerb1AsCompleted(chatId, Config.TABLE_FILE)
        }
    }



    fun buildVerbsMessage(lines: List<String>, chatId: Long): String {
        println("aaa14 buildVerbsMessage // Вход. Исходные строки: $lines")
        val line1 = lines[0]
        val line2 = lines[1]
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
        val message = if (line2 != "#") {
            listOf(line1, line2, "", line3, line4, line5, line6, line7, line8).joinToString("\n")
        } else {
            listOf(line1, line2, line3, line4, line5, line6, line7, line8).joinToString("\n")
        }
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
                    Globals.userVerb.getOrPut(chatId) { getRandomVerb() }
                    val replacement = Globals.userVerb[chatId]!!.first
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
        return result
    }

    fun replaceHashInMessage(chatId: Long, message: String): String {
        val verbPair = Globals.userVerb[chatId] ?: throw IllegalStateException("Не найдено значение для chatId $chatId")
        val replacement = "${verbPair.first} - ${verbPair.second}\n"
        return message.replace("#", replacement)
    }

    fun getRandomVerb(): Pair<String, String> {
        println("aaa_getRandomVerb // Вход в функцию получения случайной пары глаголов")
        val filePath = "Глаголы.xlsx"
        val sheetName = "Список глаголов"
        val file = checkFileExists(filePath)
        val verbPairs = readVerbPairs(file, sheetName)
        return selectRandomVerbPair(verbPairs, sheetName)
    }

    private fun checkFileExists(filePath: String): File {
        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        return file
    }

    private fun readVerbPairs(file: File, sheetName: String): MutableList<Pair<String, String>> {
        val verbPairs = mutableListOf<Pair<String, String>>()
        WorkbookFactory.create(file).use { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("Лист $sheetName не найден")
            println("aaa_getRandomVerb Лист $sheetName найден. Читаем строки...")
            // Читаем строки с индекса 1 до 88 включительно
            for (rowIndex in 1 until 89) {
                val row = sheet.getRow(rowIndex) ?: continue
                val cell1 = row.getCell(0)
                val cell2 = row.getCell(1)
                val word1 = cell1?.toString()?.trim() ?: continue
                val word2 = cell2?.toString()?.trim() ?: continue
                if (word1.isNotEmpty() && word2.isNotEmpty()) {
                    verbPairs.add(Pair(word1, word2))
                }
            }
        }
        return verbPairs
    }

    private fun selectRandomVerbPair(verbPairs: MutableList<Pair<String, String>>, sheetName: String): Pair<String, String> {
        if (verbPairs.isEmpty()) {
            throw IllegalArgumentException("Нет глаголов в листе $sheetName")
        }
        return verbPairs.random()
    }

    fun callChangeWordsVerbs1(chatId: Long, bot: Bot) {
        println("callChangeWordsVerbs1: Смена глагола для блока Глаголы 1")
        val newVerbPair: Pair<String, String> = getRandomVerb()
        Globals.userVerb[chatId] = newVerbPair
        handleBlockVerbs1(chatId, bot, keepState = true)
    }

    fun read7LinesFromColumnWithBlueHint(filePath: String, sheetName: String, columnIndex: Int, showHint: Boolean): List<String> {
        println("read7LinesFromColumnWithBlueHint: filePath=$filePath, sheetName=$sheetName, columnIndex=$columnIndex, showHint=$showHint")
        val file = File(filePath)
        if (!file.exists()) {
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        WorkbookFactory.create(file).use { workbook ->
            val sheet = workbook.getSheet(sheetName)
                ?: throw IllegalArgumentException("Лист $sheetName не найден")
            val result = mutableListOf<String>()
            for (rowIndex in 0..7) {
                val row = sheet.getRow(rowIndex)
                val cell = row?.getCell(columnIndex)
                val cellValue = if (cell != null) processVerbsCellContent(cell, showHint) else ""
                result.add(cellValue)
            }
            return result
        }
    }

    fun processVerbsCellContent(cell: Cell, showHint: Boolean): String {
        val richText = cell.richStringCellValue as? XSSFRichTextString ?: return cell.toString()
        return if (richText.numFormattingRuns() > 0) {
            processVerbsFormattedRuns(richText, richText.string, showHint)
        } else {
            processVerbsCellWithoutRuns(cell, richText.string, showHint)
        }
    }

    fun processVerbsCellWithoutRuns(cell: Cell, text: String, showHint: Boolean): String {
        println("processVerbsCellWithoutRuns: text=\"$text\", showHint=$showHint")
        val font = getCellFont(cell)
        val color = font?.let { getFontColor(it) }
        return when {
            color == "#0000FF" && !showHint -> "#"  // если синий и подсказка не показывается – возвращаем "#"
            else -> text
        }
    }

    fun processVerbsFormattedRuns(richText: XSSFRichTextString, text: String, showHint: Boolean): String {
        println("processVerbsFormattedRuns: showHint=$showHint")
        val result = buildString {
            for (i in 0 until richText.numFormattingRuns()) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < richText.numFormattingRuns()) richText.getIndexOfFormattingRun(i + 1) else text.length
                val substring = text.substring(start, end)
                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: ""
                when (colorHex) {
                    "#0000FF" -> append(if (showHint) substring else "#")
                    else -> append(substring)
                }
            }
        }
        return result
    }

    fun getCellFont(cell: Cell): XSSFFont? {
        val workbook = cell.sheet.workbook
        if (workbook is XSSFWorkbook) {
            val fontIndex = cell.cellStyle.fontIndexAsInt
            return workbook.getFontAt(fontIndex) as? XSSFFont
        }
        return null
    }

    fun getFontColor(font: XSSFFont): String {
        val xssfColor = font.xssfColor
        if (xssfColor != null && xssfColor.rgb != null) {
            return xssfColor.rgb.joinToString(prefix = "#", separator = "") { "%02X".format(it) }
        }
        return ""
    }

    fun handleToggleHintVerbs1(chatId: Long, bot: Bot, newValue: Boolean) {
        println("handleToggleHintVerbs1: chatId=$chatId, newValue=$newValue")
        Globals.userVerbsHintVisibility[chatId] = newValue
        handleBlockVerbs1(chatId, bot, keepState = true)
    }

    fun markVerb1AsCompleted(chatId: Long, tableFile: String) {
        println("markVerb1AsCompleted: Начало выполнения для chatId=$chatId, tableFile=$tableFile")
        val file = File(tableFile)
        if (!file.exists()) {
            return
        }
        WorkbookFactory.create(file).use { workbook ->
            val sheetName = "Состояние пользователя"
            val sheet = workbook.getSheet(sheetName) ?: run {
                val newSheet = workbook.createSheet(sheetName)
                newSheet
            }
            val targetRow = sheet.rowIterator().asSequence().firstOrNull { row ->
                val cell = row.getCell(0)
                // Если ячейка числовая, получаем значение без экспоненциального формата.
                val cellText = when (cell?.cellType) {
                    CellType.NUMERIC -> cell.numericCellValue.toLong().toString()
                    else -> cell?.toString()?.trim()
                }
                cell != null && cellText == chatId.toString()
            } ?: run {
                val newRowIndex = sheet.lastRowNum + 1
                sheet.createRow(newRowIndex).apply {
                    createCell(0).setCellValue(chatId.toDouble())
                }
            }
            val cellR = targetRow.getCell(17) ?: run {
                targetRow.createCell(17).also { println("Ячейка в столбце R создана") }
            }
            cellR.setCellValue(1.0)
            val tempFile = File("$tableFile.tmp")
            FileOutputStream(tempFile).use { outputStream ->
                workbook.write(outputStream)
            }
            if (file.exists() && !file.delete()) {
                throw IllegalStateException("Не удалось удалить оригинальный файл: $tableFile")
            }
            if (!tempFile.renameTo(file)) {
                throw IllegalStateException("Не удалось переименовать временный файл: ${tempFile.path}")
            }
        }
    }
}
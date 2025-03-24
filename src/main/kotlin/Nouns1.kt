import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.addScoreForPadezh
import com.github.kotlintelegrambot.dispatcher.extractWordsFromCallback
import com.github.kotlintelegrambot.dispatcher.handleBlock
import com.github.kotlintelegrambot.dispatcher.initializeUserBlockStates
import com.github.kotlintelegrambot.dispatcher.sendFinalButtons
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
        println("callNouns1: Вызов callNouns1() для chatId = $chatId")
        println("callNouns1: Установка блока 1 для chatId = $chatId")
        Globals.userStates.remove(chatId)
        Globals.userPadezh.remove(chatId)
        Globals.userBlocks[chatId] = 1
        handleBlock1(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    fun handleBlock1(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
        println("handleBlock1: Вызов handleBlock1() для chatId = $chatId, filePath = $filePath")
        println("handleBlock1: Получены значения слов: wordUz = $wordUz, wordRus = $wordRus")

        if (Globals.userPadezh[chatId] == null) {
            println("handleBlock1: Падеж для chatId = $chatId не установлен. Необходим выбор падежа.")
            sendPadezhSelection(chatId, bot, filePath)
            return
        }

        val actualWordUz = wordUz ?: "bola"
        val actualWordRus = wordRus ?: "ребенок"
        println("handleBlock1: Используем слова actualWordUz = $actualWordUz, actualWordRus = $actualWordRus")

        Globals.userWordUz[chatId] = actualWordUz
        Globals.userWordRus[chatId] = actualWordRus

        println("handleBlock1: Отправка сообщения о состоянии для chatId = $chatId")
        sendStateMessage(chatId, bot, filePath, actualWordUz, actualWordRus)
    }

    fun callPadezh(chatId: Long, bot: Bot, data: String, callbackId: String) {
        println("callPadezh: Вызов callPadezh() для chatId = $chatId с данными: $data и callbackId: $callbackId")
        val selectedPadezh = data.removePrefix("nouns1Padezh:")
        println("callPadezh: Извлечён выбранный падеж: $selectedPadezh для chatId = $chatId")

        Globals.userPadezh[chatId] = selectedPadezh
        Globals.userStates[chatId] = 0
        Globals.userColumnOrder.remove(chatId)
        println("callPadezh: Обновлены Globals (userPadezh, userStates, userColumnOrder) для chatId = $chatId")

        // При необходимости можно раскомментировать уведомление через answerCallbackQuery
        /*
        bot.answerCallbackQuery(
            callbackQueryId = callbackId,
            text = "Падеж: $selectedPadezh\nСлово: ${Globals.userWordUz[chatId]} - ${Globals.userWordRus[chatId]}",
            showAlert = true
        )
        */

        println("callPadezh: Вызов handleBlock() для chatId = $chatId")
        handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }

    fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
        println("sendPadezhSelection: Вызов sendPadezhSelection() для chatId = $chatId, filePath = $filePath")

        // Очистка предыдущих значений
        Globals.userPadezh.remove(chatId)
        Globals.userStates.remove(chatId)

        // Обновляем состояние блоков для пользователя
        initializeUserBlockStates(chatId, filePath)

        val currentBlock = 1
        println("sendPadezhSelection: Определён текущий блок = $currentBlock для chatId = $chatId")

        val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
        if (PadezhColumns == null) {
            println("sendPadezhSelection: Ошибка - невозможно загрузить данные для блока $currentBlock для chatId = $chatId")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка: невозможно загрузить данные блока."
            )
            return
        }

        val userScores = getUserScoresForBlock(chatId, filePath, PadezhColumns)
        println("sendPadezhSelection: Получены userScores для chatId = $chatId: $userScores")

        val buttons = generatePadezhSelectionButtons(currentBlock, PadezhColumns, userScores).toMutableList()


        // Если блок 2 доступен (т.е. пользователь прошёл хотя бы один падеж блока 1), добавляем кнопку "Перейти к Существительные 2"
        println("Выполненные блоки:")
        println(Globals.userBlockCompleted[chatId])

        if (Globals.userBlockCompleted[chatId]?.first == true) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("Перейти к Существительные 2", "nouns2")))
        }

        println("sendPadezhSelection: Формирование и отправка клавиатуры для выбора падежа для chatId = $chatId")

        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Выберите падеж для изучения блока $currentBlock:",
                parseMode = "MarkdownV2",
                replyMarkup = InlineKeyboardMarkup.create(buttons)
            )
        }
    }



    fun generatePadezhSelectionButtons(
        currentBlock: Int,
        PadezhColumns: Map<String, Int>,
        userScores: Map<String, Int>
    ): List<List<InlineKeyboardButton>> {
        println("generatePadezhSelectionButtons: Формирование кнопок выбора падежей для блока $currentBlock")

        val buttons = PadezhColumns.keys.map { PadezhName ->
            val score = userScores[PadezhName] ?: 0
            println("generatePadezhSelectionButtons: Обработка падежа '$PadezhName' с баллом $score")
            InlineKeyboardButton.CallbackData("$PadezhName [$score]", "nouns1Padezh:$PadezhName")
        }.map { listOf(it) }.toMutableList()
        // Пример добавления дополнительной кнопки (раскомментировать при необходимости)
        // buttons.add(listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")))
        // if (currentBlock > 1) {
        //     buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
        // }
        // if (currentBlock < 3) {
        //     buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
        // }

        println("generatePadezhSelectionButtons: Сформировано ${buttons.size} кнопок для выбора падежей")
        return buttons
    }

    fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
        println("getPadezhColumnsForBlock: Получение колонок для блока $block")
        val result = Config.COLUMN_RANGES[block]
        if (result == null) {
            println("getPadezhColumnsForBlock: Не удалось получить данные для блока $block")
        } else {
            println("getPadezhColumnsForBlock: Получены колонки: $result")
        }
        return result
    }

    fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
        println("getUserScoresForBlock: Чтение баллов пользователя для chatId = $chatId, filePath = $filePath")

        val file = File(filePath)
        if (!file.exists()) {
            println("getUserScoresForBlock: Файл '$filePath' не существует. Возврат пустых баллов.")
            return emptyMap()
        }

        val scores = mutableMapOf<String, Int>()
        val excelManager = ExcelManager(filePath)
        excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet("Состояние пользователя")
            if (sheet == null) {
                println("getUserScoresForBlock: Лист 'Состояние пользователя' не найден.")
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
                    println("getUserScoresForBlock: Найден ряд для chatId = $chatId")
                    for ((PadezhName, colIndex) in PadezhColumns) {
                        val cell = row.getCell(colIndex)
                        val score = cell?.numericCellValue?.toInt() ?: 0
                        scores[PadezhName] = score
                        println("getUserScoresForBlock: Для падежа '$PadezhName' установлен балл $score")
                    }
                    break
                }
            }
        }

        println("getUserScoresForBlock: Итоговые баллы пользователя: $scores")
        return scores
    }

//    fun processStateAndSendMessage(
//        chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
//        selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int
//    ) {
//        println("processStateAndSendMessage: Обработка состояния для chatId = $chatId, выбранный падеж: '$selectedPadezh', текущий индекс: $currentState из ${rangesForPadezh.size}")
//
//        if (currentState >= rangesForPadezh.size) {
//            println("processStateAndSendMessage: Текущее состояние $currentState превышает количество диапазонов. Добавляем балл для падежа '$selectedPadezh' и отправляем финальные кнопки.")
//            addScoreForPadezh(chatId, selectedPadezh, filePath)
//            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
//            return
//        }
//
//        val range = rangesForPadezh[currentState]
//        println("processStateAndSendMessage: Текущий диапазон: '$range'")
//
//        val currentBlock = Globals.userBlocks[chatId] ?: 1
//        val listName = when (currentBlock) {
//            1 -> "Существительные 1"
//            2 -> "Существительные 2"
//            3 -> "Существительные 3"
//            else -> "Существительные 1"
//        }
//        println("processStateAndSendMessage: Используем лист: '$listName' для блока $currentBlock")
//
//        val messageText = try {
//            generateMessageFromRange(filePath, listName, range, wordUz, wordRus)
//        } catch (e: Exception) {
//            println("processStateAndSendMessage: Ошибка при формировании сообщения: ${e.message}")
//            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
//            return
//        }
//
//        println("processStateAndSendMessage: Сформировано сообщение: '$messageText'")
//        sendMessageOrNextStep(
//            chatId,
//            bot,
//            filePath,
//            wordUz,
//            wordRus,
//            selectedPadezh,
//            rangesForPadezh,
//            currentState,
//            messageText
//        )
//    }

    fun sendStateMessage(
        chatId: Long,
        bot: Bot,
        filePath: String,
        wordUz: String?,
        wordRus: String?
    ) {
        println("sendStateMessage: Отправка сообщения для chatId = $chatId с состоянием и выбранным падежом")

        // Получаем состояние пользователя: выбранный падеж, список диапазонов и текущее состояние.
        val state = validateUserState(chatId, bot)
        if (state == null) {
            println("sendStateMessage: Состояние пользователя не валидно для chatId = $chatId")
            return
        }
        val (selectedPadezh, rangesForPadezh, currentState) = state
        println("sendStateMessage: Валидное состояние: выбранный падеж = '$selectedPadezh', текущий индекс = $currentState, диапазоны: $rangesForPadezh")

        val currentRange = rangesForPadezh[currentState]
        println("sendStateMessage: Текущий диапазон выбран: '$currentRange'")

        // Обновляем глобальные переменные для текущего пользователя.
        updateGlobals(chatId, currentRange)

        // Генерируем текст сообщения.
        val messageText = generateMessageText(filePath, "Существительные 1", currentRange, wordUz, wordRus)
            ?: run {
                bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
                return
            }
        println("sendStateMessage: Сформированное сообщение: '$messageText'")

        if (isLastStep(currentState, rangesForPadezh)) {
            handleLastStep(chatId, bot, filePath, messageText, selectedPadezh, wordUz, wordRus)
        } else {
            handleNextStep(chatId, bot, messageText, wordUz, wordRus)
        }
    }



    private fun updateGlobals(chatId: Long, currentRange: String) {
        Globals.currentSheetName[chatId] = "Существительные 1"
        Globals.currentRange[chatId] = currentRange
        println("updateGlobals: Обновлены Globals: currentSheetName = 'Существительные 1', currentRange = '$currentRange'")
    }

    private fun generateMessageText(
        filePath: String,
        sheetName: String,
        currentRange: String,
        wordUz: String?,
        wordRus: String?
    ): String? {
        return try {
            generateMessageFromRange(filePath, sheetName, currentRange, wordUz, wordRus, showHint = false)
        } catch (e: Exception) {
            println("generateMessageText: Ошибка при генерации сообщения: ${e.message}")
            null
        }
    }

    private fun isLastStep(currentState: Int, rangesForPadezh: List<String>): Boolean {
        return currentState == rangesForPadezh.size - 1
    }

    private fun handleLastStep(
        chatId: Long,
        bot: Bot,
        filePath: String,
        messageText: String,
        selectedPadezh: String,
        wordUz: String?,
        wordRus: String?
    ) {
//        if (selectedPadezh == "Исходный") {
//            println("handleLastStep: Последний падеж достигнут ('Исходный'). Переход к Существительные 2")
//            GlobalScope.launch {
//                TelegramMessageService.updateOrSendMessage(
//                    chatId = chatId,
//                    text = messageText,
//                    replyMarkup = Keyboards.transitionToNouns2ButtonWithHintToggle(
//                        wordUz,
//                        wordRus,
//                        isHintVisible = false,
//                        blockId = "nouns1"
//                    )
//                )
//                println("handleLastStep: Сообщение с кнопкой 'Перейти к Существительные 2' отправлено для chatId = $chatId")
//            }
//            addScoreForPadezh(chatId, selectedPadezh, filePath, Globals.userBlocks[chatId] ?: 1)
//        } else {
        val nextPadezh = getNextPadezh(selectedPadezh)
        println("handleLastStep: Последний шаг. Следующий падеж: $nextPadezh")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                replyMarkup = Keyboards.nextCaseButtonWithHintToggle(
                    wordUz,
                    wordRus,
                    isHintVisible = false,
                    blockId = "nouns1",
                    nextPadezh = nextPadezh
                )
            )
            println("handleLastStep: Сообщение с кнопкой перехода к следующему падежу отправлено для chatId = $chatId")
        }
        addScoreForPadezh(chatId, selectedPadezh, filePath, Globals.userBlocks[chatId] ?: 1)
    }



    private fun handleNextStep(
        chatId: Long,
        bot: Bot,
        messageText: String,
        wordUz: String?,
        wordRus: String?
    ) {
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                replyMarkup = Keyboards.nextButtonWithHintToggle(wordUz, wordRus, isHintVisible = false, blockId = "nouns1")
            )
            println("handleNextStep: Сообщение с кнопкой 'Далее' отправлено для chatId = $chatId")
        }
    }

    fun getNextPadezh(currentPadezh: String): String {
        val padezhOrder = listOf("Именительный", "Родительный", "Винительный", "Дательный", "Местный", "Исходный")
        val currentIndex = padezhOrder.indexOf(currentPadezh)
        return if (currentIndex != -1 && currentIndex < padezhOrder.size - 1) {
            padezhOrder[currentIndex + 1]
        } else {
            currentPadezh
        }
    }



//    fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
//        println("Nouns1 sendStateMessage // Отправка сообщения по текущему состоянию и падежу")
//
//        val (selectedPadezh, rangesForPadezh, currentState) = validateUserState(chatId, bot) ?: return
//
//        processStateAndSendMessage(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState)
//    }

//    fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
//        println("sendFinalButtons: Отправка финального меню действий для chatId = $chatId")
//        GlobalScope.launch {
//            TelegramMessageService.updateOrSendMessage(
//                chatId = chatId,
//                text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
//                replyMarkup = Keyboards.finalButtons(wordUz, wordRus, Globals.userBlocks[chatId] ?: 1)
//            )
//            println("sendFinalButtons: Финальное меню отправлено для chatId = $chatId")
//        }
//    }

//    fun addScoreForPadezh(
//        chatId: Long,
//        Padezh: String,
//        filePath: String,
//        sheetName: String = "Состояние пользователя"
//    ) {
//        println("addScoreForPadezh: Добавление балла для падежа '$Padezh' для chatId = $chatId, файл = $filePath, лист = $sheetName")
//        val file = File(filePath)
//        if (!file.exists()) {
//            println("addScoreForPadezh: Файл '$filePath' не найден")
//            throw IllegalArgumentException("Файл $filePath не найден.")
//        }
//
//        val PadezhColumnIndex = when (Padezh) {
//            "Именительный" -> 1
//            "Родительный" -> 2
//            "Винительный" -> 3
//            "Дательный" -> 4
//            "Местный" -> 5
//            "Исходный" -> 6
//            else -> {
//                println("addScoreForPadezh: Ошибка - неизвестный падеж: $Padezh")
//                throw IllegalArgumentException("A5 Ошибка: Неизвестный падеж: $Padezh")
//            }
//        }
//
//        val excelManager = ExcelManager(filePath)
//        excelManager.useWorkbook { workbook ->
//            val sheet = workbook.getSheet(sheetName) ?: run {
//                println("addScoreForPadezh: Ошибка - лист '$sheetName' не найден")
//                throw IllegalArgumentException("A3 Ошибка: Лист $sheetName не найден")
//            }
//
//            for (rowIndex in 1..sheet.lastRowNum) {
//                val row = sheet.getRow(rowIndex) ?: continue
//                val idCell = row.getCell(0)
//                val currentId = when (idCell?.cellType) {
//                    CellType.NUMERIC -> idCell.numericCellValue.toLong()
//                    CellType.STRING -> idCell.stringCellValue.toLongOrNull()
//                    else -> null
//                }
//                if (currentId == chatId) {
//                    println("addScoreForPadezh: Найден ряд для chatId = $chatId на строке $rowIndex")
//                    val PadezhCell = row.getCell(PadezhColumnIndex) ?: row.createCell(PadezhColumnIndex)
//                    val currentScore = PadezhCell.numericCellValue.takeIf { it > 0 } ?: 0.0
//                    println("addScoreForPadezh: Текущий балл для '$Padezh': $currentScore. Увеличиваем на 1.")
//                    PadezhCell.setCellValue(currentScore + 1)
//                    excelManager.safelySaveWorkbook(workbook)
//                    println("addScoreForPadezh: Балл для '$Padezh' обновлён и сохранён для chatId = $chatId")
//                    return@useWorkbook
//                }
//            }
//            println("addScoreForPadezh: Ряд для chatId = $chatId не найден в листе '$sheetName'")
//        }
//    }

    fun generateMessageFromRange(
        filePath: String,
        sheetName: String,
        range: String,
        wordUz: String?,
        wordRus: String?,
        showHint: Boolean = false
    ): String {
        println("generateMessageFromRange: Генерация сообщения из диапазона '$range' на листе '$sheetName' для файла '$filePath'")
        val file = File(filePath)
        if (!file.exists()) {
            println("generateMessageFromRange: Файл '$filePath' не найден")
            throw IllegalArgumentException("Файл $filePath не найден")
        }
        val excelManager = ExcelManager(filePath)
        val rawText = excelManager.useWorkbook { workbook ->
            val sheet = workbook.getSheet(sheetName) ?: run {
                println("generateMessageFromRange: Лист '$sheetName' не найден")
                throw IllegalArgumentException("Лист $sheetName не найден")
            }
            val cells = extractCellsFromRange(sheet, range, wordUz)
            val mainText = cells.firstOrNull() ?: ""
            val hintText = cells.drop(1).joinToString("\n")
            val combinedText = if (hintText.isNotBlank()) "$mainText\n\n$hintText" else mainText
            println("generateMessageFromRange: Сформирован текст: $combinedText")
            combinedText
        }
        val finalText = if (showHint) {
            println("generateMessageFromRange: Показ подсказки включён, убираем маркеры '$$'")
            rawText.replace("\$\$", "")
        } else {
            println("generateMessageFromRange: Показ подсказки отключён, замена блоков '$$ ... $$' на '*'")
            rawText.replace(Regex("""\$\$.*?\$\$""", RegexOption.DOT_MATCHES_ALL), "$wordRus \\\\- $wordUz")
        }
        println("generateMessageFromRange: Итоговое сообщение: $finalText")
        return finalText
    }

//    fun sendMessageOrNextStep(
//        chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
//        selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int, messageText: String
//    ) {
//        println("sendMessageOrNextStep: Отправка сообщения или переход к следующему шагу для chatId = $chatId, currentState = $currentState из ${rangesForPadezh.size}")
//        println("currentState $currentState; rangesForPadezh.size ${rangesForPadezh.size}++++++++++++++++++++++++++++++")
//        if (currentState == rangesForPadezh.size - 2) {
//            // Последнее сообщение для выбранного падежа – кнопка "Далее" переходит к выбору следующего падежа
//            val nextPadezh = getNextPadezh(selectedPadezh)
//            println("sendMessageOrNextStep: Последний шаг. Следующий падеж: $nextPadezh")
//            GlobalScope.launch {
//                TelegramMessageService.updateOrSendMessage(
//                    chatId = chatId,
//                    text = messageText,
//                    replyMarkup = Keyboards.nextPadezhButton(nextPadezh)
//                )
//                println("sendMessageOrNextStep: Сообщение с кнопкой перехода к следующему падежу отправлено для chatId = $chatId")
//            }
//            val currentBlock = Globals.userBlocks[chatId] ?: 1
//            println("sendMessageOrNextStep: Добавление балла и отправка финальных кнопок для блока $currentBlock")
//            com.github.kotlintelegrambot.dispatcher.addScoreForPadezh(chatId, selectedPadezh, filePath, currentBlock)
//            com.github.kotlintelegrambot.dispatcher.sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
//        } else {
//            println("sendMessageOrNextStep: Не последний шаг. Отправляем сообщение с кнопкой 'Далее'.")
//            GlobalScope.launch {
//                TelegramMessageService.updateOrSendMessage(
//                    chatId = chatId,
//                    text = messageText,
//                    replyMarkup = Keyboards.nextButton(wordUz, wordRus)
//                )
//                println("sendMessageOrNextStep: Сообщение с кнопкой 'Далее' отправлено для chatId = $chatId")
//            }
//        }
//    }

//    private fun getNextPadezh(currentPadezh: String): String {
//        // Определяем порядок падежей для блока существительных
//        val padezhOrder = listOf("Именительный", "Родительный", "Винительный", "Дательный", "Местный", "Исходный")
//        val currentIndex = padezhOrder.indexOf(currentPadezh)
//        return if (currentIndex != -1 && currentIndex < padezhOrder.size - 1) {
//            padezhOrder[currentIndex + 1]
//        } else {
//            // Если выбран последний или не найден, остаёмся с текущим
//            currentPadezh
//        }
//    }

    fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
        println("extractCellsFromRange: Извлечение ячеек из диапазона '$range' на листе '${sheet.sheetName}' для wordUz = $wordUz")
        val (start, end) = getRangeIndices(range)
        println("extractCellsFromRange: Определены индексы диапазона: start = $start, end = $end")
        val column = range[0] - 'A'
        println("extractCellsFromRange: Определён номер столбца = $column")
        val cells = (start..end).mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex)
            if (row == null) {
                println("extractCellsFromRange: Строка $rowIndex не найдена, пропуск")
                null
            } else {
                val processed = processRowForRange(row, column, wordUz)
                println("extractCellsFromRange: Строка $rowIndex обработана, результат: $processed")
                processed
            }
        }
        println("extractCellsFromRange: Извлечено ${cells.size} ячеек")
        return cells
    }

    fun getRangeIndices(range: String): Pair<Int, Int> {
        println("getRangeIndices: Обработка диапазона '$range'")
        val parts = range.split("-")
        if (parts.size != 2) {
            println("getRangeIndices: Неверный формат диапазона: $range")
            throw IllegalArgumentException("Неверный формат диапазона: $range")
        }
        val start = parts[0].replace(Regex("[A-Z]"), "").toInt() - 1
        val end = parts[1].replace(Regex("[A-Z]"), "").toInt() - 1
        println("getRangeIndices: Получены индексы: start = $start, end = $end")
        return start to end
    }

    fun processRowForRange(row: Row, column: Int, wordUz: String?): String? {
        println("processRowForRange: Обработка строки ${row.rowNum} для столбца $column с wordUz = $wordUz")
        val cell = row.getCell(column)
        val result = cell?.let { processCellContent(it, wordUz) }
        println("processRowForRange: Результат обработки строки ${row.rowNum}: $result")
        return result
    }

    fun processCellContent(cell: Cell?, wordUz: String?): String {
        println("processCellContent: Обработка содержимого ячейки, wordUz = $wordUz")
        if (cell == null) {
            println("processCellContent: Ячейка равна null, возвращаем пустую строку")
            return ""
        }
        val richText = cell.richStringCellValue as XSSFRichTextString
        val text = richText.string
        println("processCellContent: Исходный текст ячейки: $text")
        val runs = richText.numFormattingRuns()
        println("processCellContent: Количество форматированных участков: $runs")
        val result = if (runs == 0) {
            processCellWithoutRuns(cell, text, wordUz)
        } else {
            processFormattedRuns(richText, text, wordUz)
        }
        println("processCellContent: Итоговый результат: $result")
        return result
    }

    fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
        println("processCellWithoutRuns: Обработка ячейки без форматированных участков, wordUz = $wordUz")
        val font = getCellFont(cell)
        val color = font?.let { getFontColor(it) }
        //println("processCellWithoutRuns: Определён цвет шрифта: $color")
        val rawContent = adjustWordUz(text, wordUz)
        //println("processCellWithoutRuns: Текст после корректировки: $rawContent")
        val markedContent = when (color) {
            "#FF0000" -> "||$rawContent||"
            "#0000FF" -> "\$\$${rawContent}\$\$"
            else -> rawContent
        }
        //println("processCellWithoutRuns: Текст с добавленными маркерами: $markedContent")
        val escapedContent = markedContent.escapeMarkdownV2()
        //println("processCellWithoutRuns: Экранированный текст: $escapedContent")
        return escapedContent
    }

//    fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
//        println("sendWordMessage: Отправка клавиатуры для выбора слов для chatId = $chatId, файл = $filePath")
//        if (!File(filePath).exists()) {
//            println("sendWordMessage: Файл '$filePath' не найден")
//            bot.sendMessage(
//                chatId = ChatId.fromId(chatId),
//                text = "Ошибка: файл с данными не найден."
//            )
//            return
//        }
//        val inlineKeyboard = try {
//            createWordSelectionKeyboardFromExcel(filePath, "Существительные")
//        } catch (e: Exception) {
//            println("sendWordMessage: Ошибка при создании клавиатуры: ${e.message}")
//            bot.sendMessage(
//                chatId = ChatId.fromId(chatId),
//                text = "Ошибка при обработке данных: ${e.message}"
//            )
//            return
//        }
//        GlobalScope.launch {
//            TelegramMessageService.updateOrSendMessage(
//                chatId = chatId,
//                text = "Выберите слово из списка:",
//                replyMarkup = inlineKeyboard
//            )
//            println("sendWordMessage: Сообщение с выбором слова отправлено для chatId = $chatId")
//        }
//    }

//    fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
//        println("createWordSelectionKeyboardFromExcel: Создание клавиатуры из Excel-файла, файл = $filePath, лист = $sheetName")
//        val file = File(filePath)
//        if (!file.exists()) {
//            println("createWordSelectionKeyboardFromExcel: Файл '$filePath' не найден")
//            throw IllegalArgumentException("Файл $filePath не найден")
//        }
//        val excelManager = ExcelManager(filePath)
//        val buttons = excelManager.useWorkbook { workbook ->
//            val sheet = workbook.getSheet(sheetName) ?: run {
//                println("createWordSelectionKeyboardFromExcel: Лист '$sheetName' не найден")
//                throw IllegalArgumentException("D5 Ошибка: Лист $sheetName не найден")
//            }
//            val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
//            println("createWordSelectionKeyboardFromExcel: Выбраны случайные строки: $randomRows")
//            val buttons = randomRows.mapNotNull { rowIndex ->
//                val row = sheet.getRow(rowIndex)
//                if (row == null) {
//                    println("createWordSelectionKeyboardFromExcel: Строка $rowIndex не найдена, пропуск")
//                    return@mapNotNull null
//                }
//                val wordUz = row.getCell(0)?.toString()?.trim()
//                val wordRus = row.getCell(1)?.toString()?.trim()
//                if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
//                    println("createWordSelectionKeyboardFromExcel: Недостаточно данных в строке $rowIndex (wordUz: '$wordUz', wordRus: '$wordRus'), пропуск")
//                    return@mapNotNull null
//                }
//                println("createWordSelectionKeyboardFromExcel: Создание кнопки для слова: $wordUz - $wordRus")
//                InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
//            }.chunked(2)
//            println("createWordSelectionKeyboardFromExcel: Создано ${buttons.size} рядов кнопок")
//            buttons
//        }
//        return InlineKeyboardMarkup.create(buttons)
//    }

    fun validateUserState(chatId: Long, bot: Bot): Triple<String, List<String>, Int>? {
        println("validateUserState: Проверка состояния пользователя для chatId = $chatId")
        val selectedPadezh = Globals.userPadezh[chatId]
        if (selectedPadezh == null) {
            println("validateUserState: Падеж не выбран для chatId = $chatId")
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
            return null
        }
        val rangesForPadezh = Config.PADEZH_RANGES[selectedPadezh]
        if (rangesForPadezh == null) {
            println("validateUserState: Диапазоны для падежа '$selectedPadezh' не найдены")
            bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: диапазоны для падежа не найдены.")
            return null
        }
        val currentState = Globals.userStates[chatId] ?: 0
        println("validateUserState: Состояние для chatId = $chatId: выбранный падеж = '$selectedPadezh', текущий индекс = $currentState, диапазоны = $rangesForPadezh")
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
            //.replace("|", "\\|")
            .replace("{", "\\{")
            .replace("}", "\\}")
            .replace(".", "\\.")
            .replace("!", "\\!")
        return escaped
    }

    // Обработка узбекского слова в контексте строки
    fun adjustWordUz(content: String, wordUz: String?): String {
        println("adjustWordUz: Начало обработки строки. content='$content', wordUz='$wordUz'")

        fun Char.isVowel() = this.lowercaseChar() in "aeiouаеёиоуыэюя"

        val result = buildString {
            var i = 0
            while (i < content.length) {
                val char = content[i]
                //println("adjustWordUz: Обработка символа '$char' на позиции $i")
                when {
                    char == '+' && i + 1 < content.length -> {
                        val nextChar = content[i + 1]
                        //println("adjustWordUz: Найден символ '+', следующий символ '$nextChar'")
                        val lastChar = wordUz?.lastOrNull()
                        //println("adjustWordUz: Последний символ wordUz: '${lastChar ?: "нет"}'")
                        val replacement = when {
                            lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> {
                                //println("adjustWordUz: Условия для замены выполнены (vowel->'s')")
                                "s"
                            }

                            lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> {
                                //println("adjustWordUz: Условия для замены выполнены (non-vowel->'i')")
                                "i"
                            }

                            else -> {
                                //println("adjustWordUz: Условия для замены не выполнены, замена пустая")
                                ""
                            }
                        }
                        append(replacement)
                        append(nextChar)
                        //println("adjustWordUz: Добавлено: '$replacement$nextChar'")
                        i++ // Пропускаем следующий символ, так как он уже обработан
                    }

                    char == '*' -> {
                        //println("adjustWordUz: Найден символ '*', подстановка wordUz: '$wordUz'")
                        append(wordUz)
                    }

                    else -> append(char)
                }
                i++
            }
        }
        //println("adjustWordUz: Итоговый результат: '$result'")
        return result
    }

    // Обработка форматированных участков текста
    fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
        println("processFormattedRuns: Начало обработки форматированных участков текста. Исходный текст: '$text', wordUz='$wordUz'")
        val result = buildString {
            val runs = richText.numFormattingRuns()
            println("processFormattedRuns: Количество форматированных участков: $runs")
            for (i in 0 until runs) {
                val start = richText.getIndexOfFormattingRun(i)
                val end = if (i + 1 < runs) richText.getIndexOfFormattingRun(i + 1) else text.length
                println("processFormattedRuns: Обработка участка $i, start=$start, end=$end")
                val substring = text.substring(start, end)
                println("processFormattedRuns: Извлечён подстрока: '$substring'")
                val font = richText.getFontOfFormattingRun(i) as? XSSFFont
                val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
                println("processFormattedRuns: Цвет шрифта для участка $i: $colorHex")
                val adjustedSubstring = adjustWordUz(substring, wordUz)
                println("processFormattedRuns: Подстрока после adjustWordUz: '$adjustedSubstring'")
                val runContent = when (colorHex) {
                    "#FF0000" -> {
                        println("processFormattedRuns: Применение красного маркера '||' к участку $i")
                        "||$adjustedSubstring||"
                    }

                    "#0000FF" -> {
                        println("processFormattedRuns: Применение синего маркера '\$\$' к участку $i")
                        "\$\$$adjustedSubstring\$\$"
                    }

                    else -> {
                        println("processFormattedRuns: Отсутствие специальных маркеров для участка $i")
                        adjustedSubstring
                    }
                }
                append(runContent)
                println("processFormattedRuns: Итог для участка $i: '$runContent'")
            }
        }
        val finalResult = result.escapeMarkdownV2()
        println("processFormattedRuns: Итоговый результат после экранирования: '$finalResult'")
        return finalResult
    }

    // Получение шрифта ячейки
    fun getCellFont(cell: Cell): XSSFFont? {
        println("getCellFont: Попытка получить шрифт ячейки")
        val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook
        if (workbook == null) {
            println("getCellFont: Не удалось преобразовать workbook к XSSFWorkbook")
            return null
        }
        val fontIndex = cell.cellStyle.fontIndexAsInt
        println("getCellFont: Индекс шрифта ячейки: $fontIndex")
        val font = workbook.getFontAt(fontIndex) as? XSSFFont
        if (font != null) {
            println("getCellFont: Шрифт успешно получен")
        } else {
            println("getCellFont: Шрифт не найден")
        }
        return font
    }

    // Функция для извлечения цвета шрифта
    fun getFontColor(font: XSSFFont): String {
        println("getFontColor: Попытка извлечь цвет шрифта")
        val xssfColor = font.xssfColor
        if (xssfColor == null) {
            println("getFontColor: xssfColor не найден, возвращаем 'Цвет не определён'")
            return "Цвет не определён"
        }
        val rgb = xssfColor.rgb
        val result = if (rgb != null) {
            val colorHex = rgb.joinToString(prefix = "#", separator = "") { "%02X".format(it) }
            println("getFontColor: Определён цвет: $colorHex")
            colorHex
        } else {
            println("getFontColor: RGB не определён, возвращаем 'Цвет не определён'")
            "Цвет не определён"
        }
        return result
    }

    // Вспомогательный метод для получения RGB цвета с учетом оттенка
    fun XSSFColor.getRgbWithTint(): ByteArray? {
        println("getRgbWithTint: Начало получения RGB цвета с учетом оттенка")
        val baseRgb = rgb
        if (baseRgb == null) {
            println("getRgbWithTint: Базовый RGB не найден.")
            return null
        }
        val tint = this.tint
        println("getRgbWithTint: Значение оттенка: $tint")
        val result = if (tint != 0.0) {
            val tinted = baseRgb.map { (it and 0xFF.toByte()) * (1 + tint) }
                .map { it.coerceIn(0.0, 255.0).toInt().toByte() }
                .toByteArray()
            println("getRgbWithTint: Применен оттенок, результирующий RGB: ${tinted.joinToString()}")
            tinted
        } else {
            println("getRgbWithTint: Оттенок равен 0, возвращаем базовый RGB: ${baseRgb.joinToString()}")
            baseRgb
        }
        return result
    }

    // Функция, выбирающая случайное слово и обновляющая сообщение:
    fun callRandomWord(chatId: Long, bot: Bot) {
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
            // Замените "Существительные" на нужное имя листа, если требуется.
            val sheet = workbook.getSheet("Существительные")
                ?: throw Exception("Лист 'Существительные' не найден")
            // Пропускаем первую строку, если она содержит заголовки
            val randomRowIndex = (1..sheet.lastRowNum).random()
            val row = sheet.getRow(randomRowIndex)
            randomWordUz = row.getCell(0)?.toString()?.trim() ?: ""
            randomWordRus = row.getCell(1)?.toString()?.trim() ?: ""
        }
        // Обновляем глобальные переменные
        Globals.userWordUz[chatId] = randomWordUz
        Globals.userWordRus[chatId] = randomWordRus
        // Вызываем обновление сообщения (например, с помощью существующей функции handleBlock)
        handleBlock(chatId, bot, filePath, randomWordUz, randomWordRus)
    }
}
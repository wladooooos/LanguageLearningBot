package com.github.kotlintelegrambot.dispatcher

import ExcelManager
import Nouns1.handleBlock1
import com.github.kotlintelegrambot.bot
import com.github.kotlintelegrambot.dispatch
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import java.io.File
import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.messageIdList.editMessageTextViaHttp
import com.github.kotlintelegrambot.dispatcher.sendStartMenu
import com.github.kotlintelegrambot.dispatcher.sendWelcomeMessage
import com.github.kotlintelegrambot.entities.KeyboardReplyMarkup
import com.github.kotlintelegrambot.keyboards.Keyboards
import jdk.internal.joptsimple.internal.Messages.message
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFColor
import kotlin.experimental.and
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch

fun main() {
    println("main main 1. Вход в функцию main: инициализация бота")
    val bot = bot {
        token = Config.BOT_TOKEN
        dispatch {
            command("start") {
                commandStart(message.chat.id, bot)
            }
            callbackQuery {
                println("main callbackQuery 16. Вход в обработчик callback-запросов")
                val chatId = callbackQuery.message?.chat?.id
                if (chatId == null) {
                    return@callbackQuery
                }
                val data = callbackQuery.data
                if (data == null) {
                    return@callbackQuery
                }
                println("\n\n!!!!!!!!!!!!!!!  $data  !!!!!!!!!!!!!!!")
                when {
                    data == "main_menu" -> callMainMenu(chatId, bot)
                    data.startsWith("word:") -> callWord(chatId, bot, data)
                    data.startsWith("next:") -> callNext(chatId, bot, data)
                    data.startsWith("repeat:") -> callRepeat(chatId, bot, data)
                    data.startsWith("next_adjective:") -> callNextAdjective(chatId, bot)
                    data == "nouns1" -> Nouns1.callNouns1(chatId, bot)
                    data.startsWith("nouns1Padezh:") -> Nouns1.callPadezh(chatId, bot, data, callbackQuery.id)
                    data == "nouns2" -> Nouns2.callNouns2(chatId, bot)
                    data == "nouns3" -> Nouns3.callNouns3(chatId, bot)
                    data == "test" -> TestBlock.callTestBlock(chatId, bot)
                    data == "adjective1" -> Adjectives1.callAdjective1(chatId, bot)
                    data == "adjective2" -> Adjectives2.callAdjective2(chatId, bot)
                    data == "verbs1" -> Verbs1.callVerbs1(chatId, bot)
                    data == "verbs2" -> Verbs2.callVerbs2(chatId, bot)
                    data == "verbs3" -> Verbs3.callVerbs3(chatId, bot)
                    data == "change_word" -> callChangeWord(chatId, bot)
                    data == "change_words_adjective1" -> Adjectives1.callChangeWordsAdjective1(chatId, bot)
                    data == "change_words_adjective2" -> Adjectives2.callChangeWordsAdjective2(chatId, bot)
                    data == "change_Padezh" -> callChangePadezh(chatId, bot)
                    data == "reset" -> callReset(chatId, bot)
                    data == "next_block" -> callNextBlock(chatId, bot)
                    data == "prev_block" -> callPrevBlock(chatId, bot)
                    data == "next_verbs1" -> callNextVerbs1(chatId, bot)
                    data == "next_verbs2" -> callNextVerbs2(chatId, bot)
                    data == "next_verbs3" -> callNextVerbs3(chatId, bot)
                }
            }
        }
    }
    bot.startPolling()
}

fun commandStart(chatId: Long, bot: Bot) {
    println("main command(start) 5. Вход в обработчик команды /start")
    Globals.userStates.remove(chatId)
    Globals.userPadezh.remove(chatId)
    Globals.userBlocks[chatId] = 1
    Globals.userBlockCompleted.remove(chatId)
    Globals.userColumnOrder.remove(chatId)
    Globals.userWordUz[chatId] = "bola"
    Globals.userWordRus[chatId] = "ребенок"
    //sendWelcomeMessage(chatId, bot)
    sendStartMenu(chatId, bot)
}

fun callMainMenu(chatId: Long, bot: Bot) {
    println("main callbackQuery main_menu 21. Сброс состояния и возврат в главное меню")
    Globals.userStates.remove(chatId)
    Globals.userPadezh.remove(chatId)
    Globals.userWords.remove(chatId)
    Globals.userBlocks.remove(chatId)
    Globals.userBlockCompleted.remove(chatId)
    Globals.userColumnOrder.remove(chatId)
    sendStartMenu(chatId, bot)
}

//fun callPadezh(chatId: Long, bot: Bot, data: String) {
//    println("main callPadezh 23. Обработка выбора падежа")
//    val selectedPadezh = data.removePrefix("Padezh:")
//    Globals.userPadezh[chatId] = selectedPadezh
//    Globals.userStates[chatId] = 0
//    Globals.userColumnOrder.remove(chatId)
//    bot.sendMessage(
//        chatId = ChatId.fromId(chatId),
//        text = """Падеж: $selectedPadezh
//                                |Слово: ${Globals.userWordUz[chatId]} - ${Globals.userWordRus[chatId]}""".trimMargin()
//    )
//    handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
//}

fun callWord(chatId: Long, bot: Bot, data: String) {
    println("main callWord 29. Обработка выбора слова")
    val result = extractWordsFromCallback(data)
    Globals.userColumnOrder.remove(chatId)
    Globals.userWordUz[chatId] = result.first
    Globals.userWordRus[chatId] = result.second
    Globals.userWords[chatId] = result
    Globals.userStates[chatId] = 0
    handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
}

fun callNext(chatId: Long, bot: Bot, data: String) {
    println("main callNext 37. Обработка перехода к следующему состоянию")
    val params = data.removePrefix("next:").split(":")
    if (params.size == 2) {
        Globals.userWordUz[chatId] = params[0]
        Globals.userWordRus[chatId] = params[1]
        val currentState = Globals.userStates[chatId] ?: 0
        Globals.userStates[chatId] = currentState + 1
        handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }
}

fun callRepeat(chatId: Long, bot: Bot, data: String) {
    println("main callRepeat 44. Обработка повторного запуска блока с тем же словом")
    val params = data.removePrefix("repeat:").split(":")
    if (params.size == 2) {
        Globals.userWordUz[chatId] = params[0]
        Globals.userWordRus[chatId] = params[1]
        Globals.userStates[chatId] = 0
        Globals.userColumnOrder.remove(chatId)
        handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }
}

fun callNextAdjective(chatId: Long, bot: Bot) {
    println("main callNext_adjective 51. Обработка перехода в блоке прилагательных")
    val currentState = Globals.userStates[chatId] ?: 0
    Globals.userStates[chatId] = currentState + 1
    handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
}

fun callChangeWord(chatId: Long, bot: Bot) {
    println("callChangeWord: Обработка изменения слова – сброс состояния")
    Globals.userStates[chatId] = 0
    Globals.userWords.remove(chatId)
    Globals.userWordUz.remove(chatId)
    Globals.userWordRus.remove(chatId)
    Globals.userColumnOrder.remove(chatId)
    sendWordMessage(chatId, bot, Config.TABLE_FILE)
}

fun callChangePadezh(chatId: Long, bot: Bot) {
    println("callChangePadezh: Обработка изменения выбранного падежа")
    Globals.userPadezh.remove(chatId)
    Globals.userColumnOrder.remove(chatId)
    sendPadezhSelection(chatId, bot, Config.TABLE_FILE)
}

fun callReset(chatId: Long, bot: Bot) {
    println("callReset: Обработка сброса состояния падежа")
    Globals.userPadezh.remove(chatId)
    Globals.userStates.remove(chatId)
    Globals.userColumnOrder.remove(chatId)
    sendPadezhSelection(chatId, bot, Config.TABLE_FILE)
}

fun callNextBlock(chatId: Long, bot: Bot) {
    println("callNextBlock: Обработка перехода к следующему блоку")
    val currentBlock = Globals.userBlocks[chatId] ?: 1
    initializeUserBlockStates(chatId, Config.TABLE_FILE)
    val blockStates = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
    Globals.userStates.remove(chatId)
    Globals.userPadezh.remove(chatId)
    Globals.userColumnOrder.remove(chatId)
    when {
        currentBlock == 1 && blockStates.first -> {
            Globals.userBlocks[chatId] = 2
            Nouns2.handleBlock2(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
        }
        currentBlock == 2 && blockStates.second -> {
            Globals.userBlocks[chatId] = 3
            Nouns3.handleBlock3(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
        }
        currentBlock == 3 && blockStates.third -> {
            Globals.userBlocks[chatId] = 4
            checkBlocksBeforeTest(chatId, bot, Config.TABLE_FILE)
        }
        else -> {
            val notCompletedBlocks = mutableListOf<String>()
            if (!blockStates.first) notCompletedBlocks.add("Блок 1")
            if (!blockStates.second) notCompletedBlocks.add("Блок 2")
            if (!blockStates.third) notCompletedBlocks.add("Блок 3")
            val messageText = "Вы не выполнили следующие блоки:\n" +
                    notCompletedBlocks.joinToString("\n") +
                    "\nПройдите их перед тестом."
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
}

fun callPrevBlock(chatId: Long, bot: Bot) {
    println("callPrevBlock: Обработка перехода к предыдущему блоку")
    val currentBlock = Globals.userBlocks[chatId] ?: 1
    Globals.userStates.remove(chatId)
    Globals.userPadezh.remove(chatId)
    if (currentBlock > 1) {
        Globals.userBlocks[chatId] = currentBlock - 1
        handleBlock(chatId, bot, Config.TABLE_FILE, Globals.userWordUz[chatId], Globals.userWordRus[chatId])
    }
}

fun callNextVerbs1(chatId: Long, bot: Bot) {
    println("callNextVerbs1: Переход к следующему состоянию блока глаголов 1")
    val currentState = Globals.userStates[chatId] ?: 0
    Globals.userStates[chatId] = currentState + 1
    Verbs1.handleBlockVerbs1(chatId, bot)
}

fun callNextVerbs2(chatId: Long, bot: Bot) {
    println("callNextVerbs2: Переход к следующему состоянию блока глаголов 2")
    val currentState = Globals.userStates[chatId] ?: 0
    Globals.userStates[chatId] = currentState + 1
    Verbs2.handleBlockVerbs2(chatId, bot)
}

fun callNextVerbs3(chatId: Long, bot: Bot) {
    println("callNextVerbs3: Переход к следующему состоянию блока глаголов 3")
    val currentState = Globals.userStates[chatId] ?: 0
    Globals.userStates[chatId] = currentState + 1
    Verbs3.handleBlockVerbs3(chatId, bot)
}

fun sendWelcomeMessage(chatId: Long, bot: Bot) {
    println("main sendWelcomeMessage 162. Отправка приветственного сообщения")
    val keyboardMarkup = KeyboardReplyMarkup(
        keyboard = Keyboards.getStartButton().keyboard,
        resizeKeyboard = true
    )
    GlobalScope.launch {
        TelegramMessageService.updateOrSendMessage(
            chatId = chatId,
            text = """Здравствуйте!
Я бот, помогающий изучать узбекский язык!""",
            replyMarkup = null // Здесь можно не передавать клавиатуру, если это просто приветствие
        )
    }
}

fun sendStartMenu(chatId: Long, bot: Bot) {
    println("main sendStartMenu 165. Отправка стартового меню")
    GlobalScope.launch {
        TelegramMessageService.updateOrSendMessage(
        chatId = chatId,
        text = "Выберите блок для работы:",
        replyMarkup = Keyboards.startMenu()
    )
    }
}

fun handleBlock(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("main handleBlock 167. определяет текущий блок пользователя (из глобального хранилища), инициализирует состояния блоков и, в зависимости от номера блока, вызывает соответствующую функцию для обработки (например, для работы с существительными, прилагательными или тестовым блоком). Если номер блока неизвестен, функция отправляет уведомление об ошибке.")
    val currentBlock = Globals.userBlocks[chatId] ?: 1
    initializeUserBlockStates(chatId, filePath)
    println("currentBlock = $currentBlock")
    when (currentBlock) {
        1 -> handleBlock1(chatId, bot, filePath, wordUz, wordRus)
        2 -> Nouns2.handleBlock2(chatId, bot, filePath, wordUz, wordRus)
        3 -> Nouns3.handleBlock3(chatId, bot, filePath, wordUz, wordRus)
        4 -> TestBlock.handleBlockTest(chatId, bot)
        5 -> Adjectives1.handleBlockAdjective1(chatId, bot)
        6 -> Adjectives2.handleBlockAdjective2(chatId, bot)
        else -> bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Неизвестный блок: $currentBlock"
        )
    }
}

fun generateReplacements(chatId: Long) {
    println("main generateReplacements 179.  извлекает пары из глобального хранилища, формирует список ключей и, если их достаточно, создаёт мапу замен для первых 9 элементов, сохраняя результат в глобальном реестре замен.")
    val userPairs = Globals.sheetColumnPairs[chatId]
    if (userPairs == null) return
    val keysList = userPairs.keys.toList()
    val replacements = mutableMapOf<Int, String>()
    keysList.take(9).forEachIndexed { index, key ->
        replacements[index + 1] = key
    }
    Globals.userReplacements[chatId] = replacements
}

fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
    println("main sendPadezhSelection 185. Отправка меню выбора падежа")
    Globals.userPadezh.remove(chatId)
    Globals.userStates.remove(chatId)
    val currentBlock = Globals.userBlocks[chatId] ?: 1
    val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
    if (PadezhColumns == null) {
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
            chatId = chatId,
            text = "Ошибка: невозможно загрузить данные блока.",
            replyMarkup = null
        )
        }
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

fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
    println("main getPadezhColumnsForBlock 194. возвращает соответствующую мапу столбцов из глобальной конфигурации (Config.COLUMN_RANGES). Если для заданного блока данные отсутствуют, возвращается null ")
    return Config.COLUMN_RANGES[block]
}

fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
    println("main getUserScoresForBlock 196. открывает Excel-файл по заданному пути, ищет на листе \"Состояние пользователя\" строку с нужным chatId и извлекает баллы для каждого падежа (из карты PadezhColumns), формируя и возвращая результирующую карту. Если файл или лист отсутствуют, возвращается пустая карта.")
    val file = File(filePath)
    if (!file.exists()) return emptyMap()
    val scores = mutableMapOf<String, Int>()
    val excelManager = ExcelManager(filePath)
    excelManager.useWorkbook { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя") ?: return@useWorkbook
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

fun generatePadezhSelectionButtons(
    currentBlock: Int,
    PadezhColumns: Map<String, Int>,
    userScores: Map<String, Int>
): List<List<InlineKeyboardButton>> {
    println("main generatePadezhSelectionButtons 207. создает кнопки для выбора падежа: для каждого падежа из карты она формирует кнопку с названием и текущим баллом пользователя, а затем добавляет кнопки навигации («⬅️ Предыдущий блок» и «➡️ Следующий блок») в зависимости от номера текущего блока.")
    val buttons = PadezhColumns.keys.map { PadezhName ->
        val score = userScores[PadezhName] ?: 0
        InlineKeyboardButton.CallbackData("$PadezhName [$score]", "Padezh:$PadezhName")
    }.map { listOf(it) }.toMutableList()
    if (currentBlock > 1) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
    }
    if (currentBlock < 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
    }
    return buttons
}

fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
    println("main sendWordMessage 214. Отправка меню выбора слова")
    if (!File(filePath).exists()) {
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
            chatId = chatId,
            text = "Ошибка: файл с данными не найден.",
            replyMarkup = null
        )
        }
        return
    }
    val inlineKeyboard = try {
        createWordSelectionKeyboardFromExcel(filePath, "Существительные")
    } catch (e: Exception) {
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
            chatId = chatId,
            text = "Ошибка при обработке данных: ${e.message}",
            replyMarkup = null
        )
        }
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


fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
    println("main createWordSelectionKeyboardFromExcel 219. принимает путь к Excel-файлу и название листа, проверяет их наличие, выбирает случайные 10 строк из указанного листа, извлекает из каждой строки слова на узбекском и русском языках, создает для каждой пары кнопки, группирует их по 2 и возвращает готовую инлайн-клавиатуру для выбора слова.")
    val file = File(filePath)
    if (!file.exists()) {
        throw IllegalArgumentException("Файл $filePath не найден")
    }
    val excelManager = ExcelManager(filePath)
    val buttons = excelManager.useWorkbook { workbook ->
        val sheet = workbook.getSheet(sheetName)
            ?: throw IllegalArgumentException("D5 Ошибка: Лист $sheetName не найден")
        val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
        randomRows.mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex) ?: return@mapNotNull null
            val wordUz = row.getCell(0)?.toString()?.trim()
            val wordRus = row.getCell(1)?.toString()?.trim()
            if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) return@mapNotNull null
            InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
        }.chunked(2)
    }
    return InlineKeyboardMarkup.create(buttons)
}

fun extractWordsFromCallback(data: String): Pair<String, String> {
    println("main extractWordsFromCallback 233. принимает строку с callback-данными, извлекает из неё часть между \"word:(\" и \")\", затем делит полученную строку по разделителю \" - \" для получения слова на узбекском и русском языках, возвращая их в виде пары.")
    val content = data.substringAfter("word:(").substringBefore(")")
    val wordUz = content.substringBefore(" - ").trim()
    val wordRus = content.substringAfter(" - ").trim()
    return wordUz to wordRus
}

fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
    println("main sendFinalButtons 237. Отправка финального меню")
    GlobalScope.launch {
        TelegramMessageService.updateOrSendMessage(
        chatId = chatId,
        text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
        replyMarkup = Keyboards.finalButtons(wordUz, wordRus, Globals.userBlocks[chatId] ?: 1)
    )
    }
}



fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, block: Int) {
    println("main addScoreForPadezh 252. обновляет счет выбранного падежа для пользователя. На основе переданного блока она определяет нужный столбец в Excel-файле, затем ищет в листе \"Состояние пользователя\" строку с данным chatId и увеличивает значение в соответствующей ячейке на 1, сохраняя изменения.")
    val columnRanges = mapOf(
        1 to mapOf("Именительный" to 1, "Родительный" to 2, "Винительный" to 3, "Дательный" to 4, "Местный" to 5, "Исходный" to 6),
        2 to mapOf("Именительный" to 7, "Родительный" to 8, "Винительный" to 9, "Дательный" to 10, "Местный" to 11, "Исходный" to 12),
        3 to mapOf("Именительный" to 13, "Родительный" to 14, "Винительный" to 15, "Дательный" to 16, "Местный" to 17, "Исходный" to 18)
    )
    val column = columnRanges[block]?.get(Padezh) ?: return
    val excelManager = ExcelManager(filePath)
    excelManager.useWorkbook { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
            ?: throw IllegalArgumentException("A3 Ошибка: Лист 'Состояние пользователя' не найден")
        var userFound = false
        for (row in sheet) {
            val idCell = row.getCell(0) ?: continue
            val idFromCell = try { idCell.numericCellValue.toLong() } catch (e: Exception) { continue }
            if (idFromCell == chatId) {
                userFound = true
                val targetCell = row.getCell(column) ?: row.createCell(column)
                val currentValue = targetCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                targetCell.setCellValue(currentValue + 1)
                excelManager.safelySaveWorkbook(workbook)
                return@useWorkbook
            }
        }
    }
}

fun XSSFColor.getRgbWithTint(): ByteArray? {
    println("main XSSFColor.getRgbWithTint 266. Вход в функцию getRgbWithTint")
    val baseRgb = rgb ?: return null
    val tint = this.tint
    val result = if (tint != 0.0) {
        val mapped = baseRgb.map { (it and 0xFF.toByte()) * (1 + tint) }
        val clamped = mapped.map { it.coerceIn(0.0, 255.0).toInt().toByte() }
        clamped.toByteArray()
    } else {
        baseRgb
    }
    return result
}

fun checkBlocksBeforeTest(chatId: Long, bot: Bot, filePath: String) {
    println("main checkBlocksBeforeTest 277. проверяет, завершены ли все три блока пользователя перед запуском теста. Сначала она инициализирует состояния блоков, затем извлекает их статусы. Если все блоки завершены, вызывается функция для тестового блока (handleBlockTest). В противном случае формируется сообщение с перечнем незавершенных блоков и отправляется пользователю вместе с кнопкой для возврата к блокам.")
    initializeUserBlockStates(chatId, filePath)
    val (block1Completed, block2Completed, block3Completed) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
    if (block1Completed && block2Completed && block3Completed) {
        TestBlock.handleBlockTest(chatId, bot)
    } else {
        val notCompletedBlocks = mutableListOf<String>()
        if (!block1Completed) notCompletedBlocks.add("Блок 1")
        if (!block2Completed) notCompletedBlocks.add("Блок 2")
        if (!block3Completed) notCompletedBlocks.add("Блок 3")
        val messageText = "Вы не выполнили следующие блоки:\n" +
                notCompletedBlocks.joinToString("\n") +
                "\nПройдите их перед тестом."
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = messageText,
                replyMarkup = Keyboards.returnToBlocksButton()
            )
        }
    }
}

fun checkUserStateBlock3(chatId: Long, filePath: String): Boolean {
    println("main checkUserStateBlock3 284. проверяет, выполнены ли условия третьего блока для пользователя с заданным chatId. Для этого она:\n" +
            "\n" +
            "Проверяет, существует ли Excel-файл по указанному пути.\n" +
            "Открывает файл с помощью ExcelManager и пытается получить лист \"Состояние пользователя 3 блок\".\n" +
            "Извлекает первую строку листа (заголовок) и ищет в ней столбец, в котором значение равно chatId.\n" +
            "Если столбец найден, функция проверяет значения в этом столбце для строк с 1 по 30. Если хотя бы одно значение равно 0 или какая-либо строка отсутствует, функция возвращает false.\n" +
            "Если все значения удовлетворяют условию (т.е. все больше нуля), функция возвращает true.")
    val file = File(filePath)
    if (!file.exists()) return false
    val excelManager = ExcelManager(filePath)
    return excelManager.useWorkbook { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя 3 блок") ?: return@useWorkbook false
        val headerRow = sheet.getRow(0) ?: return@useWorkbook false
        var userColumnIndex: Int? = null
        for (col in 0..30) {
            val cell = headerRow.getCell(col)
            if (cell != null && cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                userColumnIndex = col
                break
            }
        }
        if (userColumnIndex == null) return@useWorkbook false
        for (rowIdx in 1..30) {
            val row = sheet.getRow(rowIdx) ?: return@useWorkbook false
            val cell = row.getCell(userColumnIndex)
            val value = cell?.numericCellValue ?: 0.0
            if (value == 0.0) return@useWorkbook false
        }
        true
    }
}

fun initializeSheetColumnPairsFromFile(chatId: Long) {
    println("main initializeSheetColumnPairsFromFile 297. Инициализирует пустую карту пар для пользователя в глобальном хранилище.\n" +
            "Проверяет наличие Excel-файла (путь из Config.TABLE_FILE) и открывает его.\n" +
            "Для каждого из листов «Существительные», «Глаголы» и «Прилагательные»:\n" +
            "Считывает все непустые пары значений из первых двух столбцов.\n" +
            "Если найдено не менее 3 пар, случайным образом выбирает 3 из них.\n" +
            "Если в сумме собрано ровно 9 пар, сохраняет их в глобальное хранилище для данного пользователя.\n" +
            "В завершение вызывает функцию generateReplacements для дальнейшей обработки данных.")
    Globals.sheetColumnPairs[chatId] = mutableMapOf()
    val file = File(Config.TABLE_FILE)
    if (!file.exists()) return
    val excelManager = ExcelManager(Config.TABLE_FILE)
    val sheetNames = listOf("Существительные", "Глаголы", "Прилагательные")
    val userPairs = mutableMapOf<String, String>()
    excelManager.useWorkbook { workbook ->
        for (sheetName in sheetNames) {
            val sheet = workbook.getSheet(sheetName) ?: continue
            val candidates = mutableListOf<Pair<String, String>>()
            for (i in 0..sheet.lastRowNum) {
                val row = sheet.getRow(i) ?: continue
                val key = row.getCell(0)?.toString()?.trim() ?: ""
                val value = row.getCell(1)?.toString()?.trim() ?: ""
                if (key.isNotEmpty() && value.isNotEmpty()) {
                    candidates.add(key to value)
                }
            }
            if (candidates.size < 3) continue
            val selectedPairs = candidates.shuffled().take(3)
            for ((key, value) in selectedPairs) {
                userPairs[key] = value
            }
        }
    }
    if (userPairs.size == 9) {
        Globals.sheetColumnPairs[chatId] = userPairs
    }
    generateReplacements(chatId)
}

//Проверка состояния!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
fun initializeUserBlockStates(chatId: Long, filePath: String) {
    println("main initializeUserBlockStates 272. определяет, завершены ли этапы (блоки) 1, 2 и 3 для конкретного пользователя. Для блоков 1 и 2 она вызывает функцию checkUserState, а для блока 3 — checkUserStateBlock3, после чего сохраняет результаты (в виде тройки булевых значений) в глобальном хранилище для данного chatId.")
    val block1Completed = checkUserState(chatId, filePath, block = 1)
    val block2Completed = checkUserState(chatId, filePath, block = 2)
    val block3Completed = checkUserStateBlock3(chatId, filePath)
    Globals.userBlockCompleted[chatId] = Triple(block1Completed, block2Completed, block3Completed)
}

fun checkUserState(
    chatId: Long,
    filePath: String,
    sheetName: String = "Состояние пользователя",
    block: Int = 1
): Boolean {
    println("main checkUserState 239. проверяет, удовлетворяет ли состояние пользователя заданным условиям для конкретного блока. Для этого она:\n" +
            "\n" +
            "Определяет диапазон столбцов для блока (блок 1: столбцы 1–6, блок 2: 7–12, блок 3: 13–18).\n" +
            "Проверяет, существует ли Excel-файл с данными, и открывает лист (по умолчанию \"Состояние пользователя\").\n" +
            "Ищет в файле строку, где первый столбец соответствует chatId, и затем проверяет, что значения во всех столбцах выбранного диапазона больше нуля.\n" +
            "Если все проверки проходят, функция возвращает true, иначе – false.")
    val columnRanges = mapOf(
        1 to (1..6),
        2 to (7..12),
        3 to (13..18)
    )
    val columns = columnRanges[block] ?: return false
    val file = File(filePath)
    if (!file.exists()) return false
    val excelManager = ExcelManager(filePath)
    var userStateFound: Boolean? = null
    excelManager.useWorkbook { workbook ->
        val sheet = workbook.getSheet(sheetName) ?: return@useWorkbook
        for (row in sheet) {
            val idCell = row.getCell(0) ?: continue
            val chatIdFromCell = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                else -> null
            }
            if (chatIdFromCell == chatId) {
                userStateFound = columns.all { colIndex ->
                    val cell = row.getCell(colIndex)
                    val value = cell?.numericCellValue ?: 0.0
                    value > 0
                }
                break
            }
        }
    }
    return userStateFound ?: false
}

//fun updateOrSendMessage(
//    chatId: Long,
//    bot: Bot,
//    text: String,
//    replyMarkup: InlineKeyboardMarkup? = null
//) {
//    println("updateOrSendMessage: Вызывается для chatId=$chatId с текстом: $text")
//    val currentMessageId = Globals.userMainMessageId[chatId]
//    // Предполагаем, что replyMarkup.toString() выдаёт корректный JSON
//    val replyMarkupJson = replyMarkup?.toString()
//    if (currentMessageId != null) {
//        println("updateOrSendMessage: Найден сохранённый message_id = $currentMessageId")
//        try {
//            // Вместо стандартного метода редактирования вызываем нашу функцию для логирования ответа
//            val response = editMessageTextViaHttp(chatId, currentMessageId.toLong(), text, replyMarkupJson)
//            println("updateOrSendMessage: Редактирование прошло, ответ: $response")
//            // Можно дополнительно проверить, что в ответе ok=true, иначе сбросить id.
//            val json = org.json.JSONObject(response)
//            if (!json.getBoolean("ok")) {
//                println("updateOrSendMessage: Ответ от Telegram не успешен, отправляем новое сообщение.")
//                Globals.userMainMessageId.remove(chatId)
//            } else {
//                return
//            }
//        } catch (e: Exception) {
//            println("updateOrSendMessage: Ошибка редактирования сообщения: ${e.message}")
//            println("updateOrSendMessage: Переходим к отправке нового сообщения.")
//        }
//    } else {
//        println("updateOrSendMessage: Сохранённого message_id не найдено, отправляем новое сообщение.")
//    }
//    // Отправляем новое сообщение через HTTP
//    val newMessageId = messageIdList.sendMessageViaHttp(chatId, text, replyMarkupJson)
//    if (newMessageId != null) {
//        Globals.userMainMessageId[chatId] = newMessageId
//        println("updateOrSendMessage: Новое сообщение отправлено, сохранён message_id = $newMessageId")
//    } else {
//        println("updateOrSendMessage: Не удалось получить message_id нового сообщения!")
//    }
//}
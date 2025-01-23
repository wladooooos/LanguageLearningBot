// Main.kt

package com.github.kotlintelegrambot.dispatcher
import com.github.kotlintelegrambot.entities.keyboard.KeyboardButton
import com.github.kotlintelegrambot.bot
import com.github.kotlintelegrambot.dispatch
import com.github.kotlintelegrambot.dispatcher.callbackQuery
import com.github.kotlintelegrambot.dispatcher.command
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import kotlin.random.Random
import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.dispatcher.generateUsersButton
import com.github.kotlintelegrambot.entities.KeyboardReplyMarkup
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import com.github.kotlintelegrambot.entities.ParseMode
import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import java.io.FileOutputStream
import kotlin.experimental.and


val PadezhRanges = mapOf(
    "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7"),
    "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14"),
    "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21"),
    "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28"),
    "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35"),
    "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42")
)

val userStates = mutableMapOf<Long, Int>() //Навигация внутри падежа
val userPadezh = mutableMapOf<Long, String>() // Хранение выбранного падежа для каждого пользователя
val userWords = mutableMapOf<Long, Pair<String, String>>() // Хранение выбранного слова для каждого пользователя
val userBlocks = mutableMapOf<Long, Int>() // Хранение текущего блока для каждого пользователя
val userBlockCompleted = mutableMapOf<Long, Triple<Boolean, Boolean, Boolean>>() // Состояния блоков (пройдено или нет)
val userColumnOrder = mutableMapOf<Long, MutableList<String>>() // Для хранения случайного порядка столбцов
var wordUz: String? = null
var wordRus: String? = null
val tableFile = "Алгоритм 3.7.xlsx"

fun main() {
    println("ЖЖЖ main // Основная функция запуска бота")

    val bot = bot {
        token = "7856005284:AAFVvPnRadWhaotjUZOmFyDFgUHhZ0iGsCo"
        println("Ж1 Инициализация бота с токеном")

        dispatch {
            command("start") {
                val chatId = message.chat.id
                println("Ж2 Команда /start от пользователя: chatId = $chatId")

                // Полный сброс состояния
                userStates.remove(chatId)
                userPadezh.remove(chatId)
                userWords.remove(chatId)
                userBlocks[chatId] = 1
                userBlockCompleted.remove(chatId)
                userColumnOrder.remove(chatId)
                println("Ж3 Сброс состояния для пользователя: chatId = $chatId")

                sendWelcomeMessage(chatId, bot)
                println("Ж4 Отправлено приветственное сообщение для пользователя: chatId = $chatId")
                handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                println("Ж5 Вызвана функция handleBlock для пользователя: chatId = $chatId")
            }

            callbackQuery {
                val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery
                val data = callbackQuery.data ?: return@callbackQuery
                println("Ж6 Получен callback от пользователя: chatId = $chatId, data = $data")

                when {
                    data.startsWith("Padezh:") -> {
                        val selectedPadezh = data.removePrefix("Padezh:")
                        println("Ж7 Выбранный падеж: chatId = $chatId, selectedPadezh = $selectedPadezh")
                        userPadezh[chatId] = selectedPadezh
                        userStates[chatId] = 0
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Вы выбрали: $selectedPadezh."
                        )
                        println("Ж8 Сообщение с подтверждением падежа отправлено: chatId = $chatId, selectedPadezh = $selectedPadezh")
                        handleBlock(chatId, bot, tableFile, wordUz, wordRus)
//                        if (userWords.containsKey(chatId)) {
//                            wordUz = userWords[chatId]!!.first
//                            wordRus = userWords[chatId]!!.second
//                            println("Ж9 Повторный вызов handleBlock: chatId = $chatId, wordUz = $wordUz, wordRus = $wordRus")
//                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
//                        } else {
//                            bot.sendMessage(
//                                chatId = ChatId.fromId(chatId),
//                                text = "Теперь выберите слово."
//                            )
//                        }
                    }
                    data.startsWith("word:") -> {
                        println("Ж11 Выбор слова: data = $data")
                        val result = extractWordsFromCallback(data)
                        userColumnOrder.remove(chatId)
                        wordUz = result.first
                        wordRus = result.second
                        println("Ж12 Слово выбрано: chatId = $chatId, wordUz = $wordUz, wordRus = $wordRus")
                        userWords[chatId] = result
                        userStates[chatId] = 0
                        handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                        println("Ж13 Вызвана sendStateMessage: chatId = $chatId, wordUz = $wordUz, wordRus = $wordRus")
                    }
                    data.startsWith("next:") -> {
                        println("Ж14 Обработка кнопки 'Далее': data = $data")
                        val params = data.removePrefix("next:").split(":")
                        if (params.size == 2) {
                            wordUz = params[0]
                            wordRus = params[1]
                            val currentState = userStates[chatId] ?: 0
                            userStates[chatId] = currentState + 1
                            println("Ж15 Состояние обновлено: chatId = $chatId, currentState = ${userStates[chatId]}")

                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                            println("Ж16 Вызвана handleBlock для кнопки 'Далее': chatId = $chatId, wordUz = $wordUz, wordRus = $wordRus")
                        }
                    }
                    data.startsWith("repeat:") -> {
                        println("Ж17 Обработка кнопки 'Повторить': data = $data")
                        val params = data.removePrefix("repeat:").split(":")
                        if (params.size == 2) {
                            wordUz = params[0]
                            wordRus = params[1]
                            userStates[chatId] = 0
                            userColumnOrder.remove(chatId)
                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                            println("Ж18 Вызвана sendStateMessage для кнопки 'Повторить': chatId = $chatId, wordUz = $wordUz, wordRus = $wordRus")
                        }
                    }
                    data == "change_word" -> {
                        println("Ж19 Выбор нового слова для пользователя: chatId = $chatId")
                        userStates[chatId] = 0
                        userWords.remove(chatId)
                        wordUz = null
                        wordRus = null
                        userColumnOrder.remove(chatId)
                        sendWordMessage(chatId, bot, tableFile)
                        println("Ж20 Отправлена клавиатура для выбора слов: chatId = $chatId")
                    }
                    data == "change_Padezh" -> {
                        println("Ж21 Выбор нового падежа для пользователя: chatId = $chatId")
                        userPadezh.remove(chatId)
                        userColumnOrder.remove(chatId)
                        sendPadezhSelection(chatId, bot, tableFile)
                        println("Ж22 Отправлена клавиатура для выбора падежей: chatId = $chatId")
                    }
                    data == "reset" -> {
                        println("Ж23 Полный сброс данных для пользователя: chatId = $chatId")
                        userWords.remove(chatId)
                        userPadezh.remove(chatId)
                        userStates.remove(chatId)
                        sendPadezhSelection(chatId, bot, tableFile)
                        println("Ж24 Сброс данных завершен: chatId = $chatId")
                    }
                    data == "next_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        val blockStates = userBlockCompleted[chatId] ?: Triple(false, false, false)
                        println("Ж25 Обработка 'Следующий блок': chatId = $chatId, currentBlock = $currentBlock")
                        userStates.remove(chatId)
                        userPadezh.remove(chatId)
                        userWords.remove(chatId)
                        userWords.remove(chatId)
                        wordUz = null
                        wordRus = null

                        if ((currentBlock == 1 && blockStates.first) ||
                            (currentBlock == 2 && blockStates.second) ||
                            currentBlock == 3
                        ) {
                            if (currentBlock < 3) {
                                userBlocks[chatId] = currentBlock + 1
                                handleBlock(chatId, bot, tableFile, wordUz, wordRus)
//                                println("Ж26 Переключение на следующий блок: chatId = $chatId, nextBlock = ${userBlocks[chatId]}")
                                //initializeUserBlockStates(chatId, tableFile)
//                                sendPadezhSelection(chatId, bot, tableFile)
                            } else {
                                bot.sendMessage(
                                    chatId = ChatId.fromId(chatId),
                                    text = "Вы уже на последнем блоке."
                                )
                                println("Ж27 Пользователь на последнем блоке: chatId = $chatId")
                            }
                        } else {
                            bot.sendMessage(
                                chatId = ChatId.fromId(chatId),
                                text = "Пройдите все падежи текущего блока, чтобы открыть следующий."
                            )
                            println("Ж28 Блок $currentBlock не завершен: chatId = $chatId")
                        }
                    }
                    data == "prev_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        println("Ж29 Обработка 'Предыдущий блок': chatId = $chatId, currentBlock = $currentBlock")
                        userStates.remove(chatId)
                        userPadezh.remove(chatId)
                        userWords.remove(chatId)
                        userWords.remove(chatId)
                        wordUz = null
                        wordRus = null
                        if (currentBlock > 1) {
                            userBlocks[chatId] = currentBlock - 1
                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                            println("Ж30 Возврат на предыдущий блок: chatId = $chatId, prevBlock = ${userBlocks[chatId]}")
                        }
                    }
                }
            }
        }
    }

    bot.startPolling()
    println("Ж31 Бот начал опрос обновлений")
}

fun handleBlock(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("AAA handleBlock // Обработка основного блока для пользователя")
    println("A1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")

    val currentBlock = userBlocks[chatId] ?: 1
    println("A2 Текущий блок пользователя: $currentBlock, userPadezh=${userPadezh[chatId]}")

    initializeUserBlockStates(chatId, filePath)
    println("A3 Состояния блоков пользователя после инициализации: ${userBlockCompleted[chatId]}")

    if (userPadezh[chatId] == null) {
        println("A4 Падеж не выбран для пользователя. Отправляем выбор падежа.")
        sendPadezhSelection(chatId, bot, filePath)
        return
    }
    println("A5 Падеж выбран: ${userPadezh[chatId]}")

    when (currentBlock) {
        1 -> {
            println("A6 Запуск handleBlock1 для пользователя $chatId")
            handleBlock1(chatId, bot, filePath, wordUz, wordRus)
        }
        2 -> {
            println("A7 Запуск handleBlock2 для пользователя $chatId")
            handleBlock2(chatId, bot, filePath, wordUz, wordRus)
        }
        3 -> {
            println("A8 Запуск handleBlock3 для пользователя $chatId")
            handleBlock3(chatId, bot, filePath, wordUz, wordRus)
        }
        else -> {
            println("A9 Неизвестный блок: $currentBlock для пользователя $chatId")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Неизвестный блок: $currentBlock"
            )
        }
    }
    println("A10 Выход из функции handleBlock")
}

fun handleBlock1(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("BBB handleBlock1 // Обработка блока 1")
    if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
        sendWordMessage(chatId, bot, tableFile)
        println("B0 Отправлена клавиатура для выбора слов: chatId = $chatId")
    } else {
        println("B1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")
        sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
        println("B2 Выход из функции handleBlock1")
    }
}

fun handleBlock2(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("CCC handleBlock2 // Обработка блока 2")
    println("C1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")

    val selectedPadezh = userPadezh[chatId]
    println("C2 Выбранный падеж: $selectedPadezh")
    if (selectedPadezh == null) {
        println("C3 Падеж не выбран. Сообщаем пользователю об ошибке.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
        return
    }

    if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
        println("Слово отсутствует. Запускаем выбор слова.")
        sendWordMessage(chatId, bot, tableFile)
    } else {

    val blockRanges = mapOf(
        "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7"),
        "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14"),
        "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21"),
        "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28"),
        "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"),
        "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42", "F36-F42")
    )[selectedPadezh] ?: return
    println("C4 Диапазоны для выбранного падежа: $blockRanges")

    if (userColumnOrder[chatId].isNullOrEmpty()) {
        userColumnOrder[chatId] = blockRanges.shuffled().toMutableList()
        println("C5 Новый перемешанный порядок столбцов: ${userColumnOrder[chatId]}")
    }

    val currentState = userStates[chatId] ?: 0
    val shuffledColumns = userColumnOrder[chatId]!!
    println("C6 Текущее состояние: $currentState, Перемешанный список столбцов: $shuffledColumns")

    if (currentState >= shuffledColumns.size) {
        println("C7 Все столбцы блока 2 завершены. Обновляем баллы.")
        addScoreForPadezh(chatId, selectedPadezh, filePath, block = 2)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    println("C8 Обработка столбцов от $currentState до ${shuffledColumns.size - 1}")
    for (i in currentState until shuffledColumns.size) {
        val range = shuffledColumns[i]
        println("C9 Текущий диапазон для обработки: $range")

        val messageText = generateMessageFromRange(filePath, "Существительные 2", range, wordUz, wordRus)
        println("C10 Сформированное сообщение для диапазона: $messageText")

        val isLastMessage = i == shuffledColumns.size - 1
        println("C11 Последний столбец? $isLastMessage")

        if (isLastMessage) {
            println("C12 Отправляем последнее сообщение без кнопки.")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = messageText,
                parseMode = ParseMode.MARKDOWN_V2
            )
            addScoreForPadezh(chatId, selectedPadezh, filePath, block = 2)
            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        } else {
            println("C13 Отправляем сообщение с кнопкой 'Далее'.")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = messageText,
                parseMode = ParseMode.MARKDOWN_V2,
                replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                    InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus")
                )
            )
        }

        if (!isLastMessage) {
            println("C14 Ожидание действия пользователя.")
            return
        }
    }
    }
    println("C15 Завершение обработки блока 2.")
}

fun handleBlock3(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("DDD handleBlock3 // Обработка блока 3")
    println("D1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")
    sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
    println("D2 Выход из функции handleBlock3")
}

fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
    println("EEE sendPadezhSelection // Формирование клавиатуры для выбора падежа")
    println("E1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath")

    userPadezh.remove(chatId)
    userStates.remove(chatId)
    println("E2 Сброс текущего выбранного падежа и состояния для пользователя $chatId")

    val currentBlock = userBlocks[chatId] ?: 1
    println("E3 Текущий блок пользователя: $currentBlock")

    val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
    if (PadezhColumns == null) {
        println("E4 Ошибка: блок $currentBlock не найден.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: невозможно загрузить данные блока.")
        return
    }

    val userScores = getUserScoresForBlock(chatId, filePath, PadezhColumns)
    println("E5 Полученные баллы пользователя: $userScores")

    val buttons = generatePadezhSelectionButtons(currentBlock, PadezhColumns, userScores)
    println("E6 Сформирована клавиатура с кнопками: $buttons")

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите падеж для изучения блока $currentBlock:",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
    println("E7 Сообщение с клавиатурой отправлено пользователю $chatId")
}

fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
    println("FFF getPadezhColumnsForBlock // Получение колонок текущего блока")
    println("F1 Вход в функцию. Параметры: block=$block")
    val columnRanges = mapOf(
        1 to mapOf("Именительный" to 1, "Родительный" to 2, "Винительный" to 3, "Дательный" to 4, "Местный" to 5, "Исходный" to 6),
        2 to mapOf("Именительный" to 7, "Родительный" to 8, "Винительный" to 9, "Дательный" to 10, "Местный" to 11, "Исходный" to 12),
        3 to mapOf("Именительный" to 13, "Родительный" to 14, "Винительный" to 15, "Дательный" to 16, "Местный" to 17, "Исходный" to 18)
    )
    val result = columnRanges[block]
    println("F2 Результат: $result")
    return result
}

fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
    println("GGG getUserScoresForBlock // Чтение баллов пользователя для блока")
    println("G1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, PadezhColumns=$PadezhColumns")

    val file = File(filePath)
    if (!file.exists()) {
        println("G2 Ошибка: файл $filePath не найден.")
        return emptyMap()
    }

    val scores = mutableMapOf<String, Int>()
    println("G3 Открываем файл для чтения данных пользователя.")
    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
        if (sheet == null) {
            println("G4 Ошибка: лист 'Состояние пользователя' не найден.")
            return emptyMap()
        }

        for (row in sheet) {
            val idCell = row.getCell(0)
            val chatIdXlsx = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                else -> null
            }

            if (chatId == chatIdXlsx) {
                println("G5 Найдена строка пользователя $chatIdXlsx. Извлекаем баллы...")
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

    println("G7 Итоговые баллы для пользователя $chatId: $scores")
    return scores
}

fun generatePadezhSelectionButtons(
    currentBlock: Int,
    PadezhColumns: Map<String, Int>,
    userScores: Map<String, Int>
): List<List<InlineKeyboardButton>> {
    println("HHH generatePadezhSelectionButtons // Формирование кнопок выбора падежей")
    println("H1 Вход в функцию. Параметры: currentBlock=$currentBlock, PadezhColumns=$PadezhColumns, userScores=$userScores")

    val buttons = PadezhColumns.keys.map { PadezhName ->
        val score = userScores[PadezhName] ?: 0
        InlineKeyboardButton.CallbackData("$PadezhName [$score]", "Padezh:$PadezhName")
    }.map { listOf(it) }.toMutableList()
    println("H2 Кнопки для выбора падежей: $buttons")

    if (currentBlock > 1) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
        println("H3 Добавлена кнопка для перехода к предыдущему блоку")
    }
    if (currentBlock < 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
        println("H4 Добавлена кнопка для перехода к следующему блоку")
    }

    println("H5 Итоговые кнопки: $buttons")
    return buttons
}

fun sendWelcomeMessage(chatId: Long, bot: Bot) {
    println("III sendWelcomeMessage // Отправка приветственного сообщения")
    println("I1 Вход в функцию. Параметры: chatId=$chatId")

    val keyboardMarkup = KeyboardReplyMarkup(
        keyboard = generateUsersButton(),
        resizeKeyboard = true
    )
    println("I2 Сформирована клавиатура для пользователя $chatId")

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = """Здравствуйте!
Я бот, помогающий изучать узбекский язык!""",
        replyMarkup = keyboardMarkup
    )
    println("I3 Сообщение отправлено пользователю $chatId")
}

fun generateUsersButton(): List<List<KeyboardButton>> {
    println("JJJ generateUsersButton // Генерация кнопки /start")
    println("J1 Вход в функцию.")
    val result = listOf(
        listOf(KeyboardButton("/start"))
    )
    println("J2 Сформированы кнопки: $result")
    return result
}

fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
    println("KKK sendWordMessage // Отправка клавиатуры с выбором слов")
    println("K1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath")

    if (!File(filePath).exists()) {
        println("K2 Ошибка: файл $filePath не найден.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка: файл с данными не найден."
        )
        println("K3 Сообщение об ошибке отправлено пользователю $chatId")
        return
    }

    val inlineKeyboard = try {
        println("K4 Генерация клавиатуры из файла $filePath")
        createWordSelectionKeyboardFromExcel(filePath, "Существительные")
    } catch (e: Exception) {
        println("K5 Ошибка при обработке данных: ${e.message}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка при обработке данных: ${e.message}"
        )
        println("K6 Сообщение об ошибке отправлено пользователю $chatId")
        return
    }

    println("K7 Клавиатура успешно создана: $inlineKeyboard")
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите слово из списка:",
        replyMarkup = inlineKeyboard
    )
    println("K8 Сообщение с клавиатурой отправлено пользователю $chatId")
}

fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
    println("LLL createWordSelectionKeyboardFromExcel // Создание клавиатуры из Excel-файла")
    println("L1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName")

    println("L2 Проверка существования файла $filePath")
    val file = File(filePath)
    if (!file.exists()) {
        println("L3 Ошибка: файл $filePath не найден")
        throw IllegalArgumentException("Файл $filePath не найден")
    }

    println("L4 Файл найден: $filePath. Открываем файл Excel")
    val workbook = WorkbookFactory.create(file)
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
    workbook.close()
    println("L14 Файл Excel закрыт. Генерация клавиатуры завершена")

    return InlineKeyboardMarkup.create(buttons)
}

fun extractWordsFromCallback(data: String): Pair<String, String> {
    println("MMM extractWordsFromCallback // Извлечение слов из callback data")
    println("M1 Вход в функцию. Параметры: data=$data")

    val content = data.substringAfter("word:(").substringBefore(")")
    val wordUz = content.substringBefore(" - ").trim()
    val wordRus = content.substringAfter(" - ").trim()

    println("M2 Извлеченные слова: wordUz=$wordUz, wordRus=$wordRus")
    return wordUz to wordRus
}

fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("NNN sendStateMessage // Отправка сообщения по текущему состоянию и падежу")
    println("N1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")

    val (selectedPadezh, rangesForPadezh, currentState) = validateUserState(chatId, bot) ?: return
    println("N2 Валидное состояние: selectedPadezh=$selectedPadezh, rangesForPadezh=$rangesForPadezh, currentState=$currentState")

    processStateAndSendMessage(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState)
    println("N3 Завершение sendStateMessage")
}

fun validateUserState(chatId: Long, bot: Bot): Triple<String, List<String>, Int>? {
    println("OOO validateUserState // Проверка состояния пользователя")
    println("O1 Вход в функцию. Параметры: chatId=$chatId")

    val selectedPadezh = userPadezh[chatId]
    if (selectedPadezh == null) {
        println("O2 Ошибка: выбранный падеж отсутствует.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
        return null
    }
    println("O3 Выбранный падеж: $selectedPadezh")

    val rangesForPadezh = PadezhRanges[selectedPadezh]
    if (rangesForPadezh == null) {
        println("O4 Ошибка: диапазоны для падежа $selectedPadezh отсутствуют.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: диапазоны для падежа не найдены.")
        return null
    }

    val currentState = userStates[chatId] ?: 0
    println("O5 Текущее состояние пользователя: $currentState")

    return Triple(selectedPadezh, rangesForPadezh, currentState)
}


fun processStateAndSendMessage(
    chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
    selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int
) {
    println("PPP processStateAndSendMessage // Обработка состояния и отправка сообщения")
    println("P1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus, selectedPadezh=$selectedPadezh, currentState=$currentState")

    if (currentState >= rangesForPadezh.size) {
        println("P2 Все этапы завершены для падежа: $selectedPadezh")
        addScoreForPadezh(chatId, selectedPadezh, filePath)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    val range = rangesForPadezh[currentState]
    println("P3 Генерация сообщения для диапазона: $range")
    val currentBlock = userBlocks[chatId] ?: 1
    val listName = when (currentBlock) {
        1 -> "Существительные 1"
        2 -> "Существительные 2"
        3 -> "Существительные 3"
        else -> "Существительные 1"
    }
    println("P4 Текущий блок: $currentBlock, Лист: $listName")

    val messageText = try {
        generateMessageFromRange(filePath, listName, range, wordUz, wordRus)
    } catch (e: Exception) {
        println("P5 Ошибка при генерации сообщения: ${e.message}")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        return
    }

    println("P6 Сообщение сгенерировано: $messageText")
    sendMessageOrNextStep(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState, messageText)
    println("P7 Завершение processStateAndSendMessage")
}

fun sendMessageOrNextStep(
    chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
    selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int, messageText: String
) {
    println("QQQ sendMessageOrNextStep // Отправка сообщения или переход к следующему шагу")
    println("Q1 Вход в функцию. Параметры: chatId=$chatId, currentState=$currentState, messageText=$messageText")

    if (currentState == rangesForPadezh.size - 1) {
        println("Q2 Последний этап для текущего падежа: $selectedPadezh")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2
        )
        println("Q3 Сообщение отправлено. Добавляем балл и отправляем финальное меню.")
        val currentBlock = userBlocks[chatId] ?: 1
        addScoreForPadezh(chatId, selectedPadezh, filePath, currentBlock)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        println("Q4 Финальное меню отправлено.")
    } else {
        println("Q5 Этап не последний. Отправляем сообщение с кнопкой 'Далее'")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2,
            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus")
            )
        )
        println("Q6 Сообщение отправлено.")
    }
}

fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
    println("RRR sendFinalButtons // Отправка финального меню действий")
    println("R1 Вход в функцию. Параметры: chatId=$chatId, wordUz=$wordUz, wordRus=$wordRus")

    val showNextStep = checkUserState(chatId, filePath)
    println("R2 Проверка состояния пользователя. Добавить кнопку 'Следующий блок': $showNextStep")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Повторить", "repeat:$wordUz:$wordRus")),
        listOf(InlineKeyboardButton.CallbackData("Изменить слово", "change_word")),
        listOf(InlineKeyboardButton.CallbackData("Изменить падеж", "change_Padezh"))
    )

    val currentBlock = userBlocks[chatId] ?: 1
    if (currentBlock > 1) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
    }
    if (currentBlock < 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
    }

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
    println("R3 Финальное меню отправлено пользователю $chatId")
}

fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?): String {
    println("SSS generateMessageFromRange // Генерация сообщения из диапазона Excel")
    println("S1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName, range=$range, wordUz=$wordUz, wordRus=$wordRus")

    val file = File(filePath)
    if (!file.exists()) {
        println("S2 Ошибка: файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден")
    }

    val workbook = WorkbookFactory.create(file)
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

    workbook.close()
    println("S6 Файл Excel закрыт.")

    val result = listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
    println("S7 Генерация завершена. Результат: $result")
    return result
}

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

fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
    println("VVV extractCellsFromRange // Извлечение и обработка ячеек из диапазона")
    println("V1 Вход в функцию. Параметры: range=\"$range\", wordUz=\"$wordUz\"")

    val (start, end) = range.split("-").map { it.replace(Regex("[A-Z]"), "").toInt() - 1 }
    val column = range[0] - 'A'
    println("V2 Диапазон строк: $start-$end, Колонка: $column")

    val cells = (start..end).mapNotNull { rowIndex ->
        val row = sheet.getRow(rowIndex)
        if (row == null) {
            println("V3 ⚠️ Строка $rowIndex отсутствует, пропускаем.")
            return@mapNotNull null
        }
        val cell = row.getCell(column)
        if (cell == null) {
            println("V4 ⚠️ Ячейка в строке $rowIndex отсутствует, пропускаем.")
            return@mapNotNull null
        }
        val processed = processCellContent(cell, wordUz)
        println("V5 ✅ Обработанная ячейка в строке $rowIndex: \"$processed\"")
        processed
    }
    println("V6 Результат извлечения: $cells")
    return cells
}

fun checkUserState(chatId: Long, filePath: String, sheetName: String = "Состояние пользователя"): Boolean {
    println("WWW checkUserState // Проверка состояния пользователя")
    println("W1 Вход в функцию. Параметры: chatId=$chatId, filePath=\"$filePath\", sheetName=\"$sheetName\"")

    val file = File(filePath)
    if (!file.exists()) {
        println("W2 ❌ Файл $filePath не найден.")
        return false
    }

    val allCompleted: Boolean
    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            println("W3 ❌ Лист $sheetName не найден.")
            return false
        }
        println("W4 ✅ Лист найден. Проверка строк...")

        var userRow: Row? = null
        for (rowIndex in 1..sheet.lastRowNum.coerceAtLeast(1)) {
            val row = sheet.getRow(rowIndex) ?: sheet.createRow(rowIndex)
            val idCell = row.getCell(0) ?: row.createCell(0)

            val currentId = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                else -> null
            }
            println("W5 Проверяем строку ${rowIndex + 1}. ID: $currentId")

            if (currentId == chatId) {
                userRow = row
                println("W6 ✅ Пользователь найден в строке ${rowIndex + 1}")
                break
            }

            if (currentId == null || currentId == 0L) {
                println("W7 ⚠️ Пустая строка. Добавляем нового пользователя.")
                idCell.setCellValue(chatId.toDouble())
                for (i in 1..6) row.createCell(i).setCellValue(0.0)
                safelySaveWorkbook(workbook, filePath)
                println("C8 ✅ Новый пользователь добавлен в строку ${rowIndex + 1}")
                return false
            }
        }

        allCompleted = userRow?.let {
            (1..6).all { index ->
                val cell = it.getCell(index)
                val value = cell?.toString()?.toDoubleOrNull() ?: 0.0
                println("W9 Проверка колонки $index. Значение: $value")
                value > 0
            }
        } ?: false
    }
    println("W10 Все этапы завершены: $allCompleted")
    return allCompleted
}

fun safelySaveWorkbook(workbook: org.apache.poi.ss.usermodel.Workbook, filePath: String) {
    println("XXX safelySaveWorkbook // Сохранение данных в файл")
    println("X1 Вход в функцию. Параметры: filePath=\"$filePath\"")

    val tempFile = File("${filePath}.tmp")
    try {
        FileOutputStream(tempFile).use { workbook.write(it) }
        println("X2 Данные успешно сохранены во временный файл: ${tempFile.path}")

        val originalFile = File(filePath)
        if (originalFile.exists() && !originalFile.delete()) {
            println("X3 ❌ Ошибка: не удалось удалить оригинальный файл.")
            throw IllegalStateException("Не удалось удалить файл: $filePath")
        }

        if (!tempFile.renameTo(originalFile)) {
            println("X4 ❌ Ошибка: не удалось переименовать временный файл.")
            throw IllegalStateException("Не удалось переименовать временный файл: ${tempFile.path}")
        }
        println("X5 ✅ Файл успешно сохранен: $filePath")
    } catch (e: Exception) {
        println("X6 ❌ Ошибка при сохранении файла: ${e.message}")
        tempFile.delete()
        throw e
    }
}




fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, sheetName: String = "Состояние пользователя") {
    println("YYY addScoreForPadezh // Добавляет балл для падежа")

    println("Y1 Начало работы. Пользователь: $chatId, Падеж: $Padezh, Файл: $filePath, Лист: $sheetName")
    val file = File(filePath)
    if (!file.exists()) {
        println("Y2 Ошибка: Файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден.")
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
            ?: throw IllegalArgumentException("A3 Ошибка: Лист $sheetName не найден")
        println("Y4 Лист $sheetName найден")

        val PadezhColumnIndex = when (Padezh) {
            "Именительный" -> 1
            "Родительный" -> 2
            "Винительный" -> 3
            "Дательный" -> 4
            "Местный" -> 5
            "Исходный" -> 6
            else -> throw IllegalArgumentException("A5 Ошибка: Неизвестный падеж: $Padezh")
        }

        println("Y6 Поиск пользователя $chatId в таблице...")
        for (rowIndex in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue
            val idCell = row.getCell(0)
            val currentId = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            println("Y7 Проверяем строку ${rowIndex + 1}: ID = $currentId")
            if (currentId == chatId) {
                println("Y8 Пользователь найден в строке ${rowIndex + 1}")
                val PadezhCell = row.getCell(PadezhColumnIndex) ?: row.createCell(PadezhColumnIndex)
                val currentScore = PadezhCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("Y9 Текущее значение: $currentScore")
                PadezhCell.setCellValue(currentScore + 1)
                println("Y10 Новое значение: ${currentScore + 1}")
                safelySaveWorkbook(workbook, filePath)
                println("Y11 Балл добавлен и изменения сохранены")
                return
            }
        }
        println("Y12 Ошибка: Пользователь $chatId не найден.")
    }
}


fun checkUserState(chatId: Long, filePath: String, sheetName: String = "Состояние пользователя", block: Int = 1): Boolean {
    println("ZZZ checkUserState // Проверка состояния пользователя")

    println("Z1 Начало проверки. Пользователь: $chatId, Файл: $filePath, Лист: $sheetName, Блок: $block")
    val columnRanges = mapOf(
        1 to (1..6),
        2 to (7..12),
        3 to (13..18)
    )
    val columns = columnRanges[block] ?: run {
        println("Z2 Ошибка: Неизвестный блок: $block")
        return false
    }

    println("Z3 Диапазон колонок для блока $block: $columns")
    val file = File(filePath)
    if (!file.exists()) {
        println("Z4 Ошибка: Файл $filePath не найден")
        return false
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            println("Z5 Ошибка: Лист $sheetName не найден")
            return false
        }

        println("Z6 Лист $sheetName найден. Всего строк: ${sheet.lastRowNum + 1}")
        for (row in sheet) {
            val idCell = row.getCell(0) ?: continue
            val chatIdFromCell = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            println("Z7 Проверяем строку ${row.rowNum + 1}: ID = $chatIdFromCell")
            if (chatIdFromCell == chatId) {
                println("Z8 Пользователь найден в строке ${row.rowNum + 1}")
                val allColumnsCompleted = columns.all { colIndex ->
                    val cell = row.getCell(colIndex)
                    val value = cell?.numericCellValue ?: 0.0
                    println("Z9 Колонка $colIndex: значение = $value. Выполнено? ${value > 0}")
                    value > 0
                }
                println("Z10 Результат проверки: $allColumnsCompleted")
                return allColumnsCompleted
            }
        }
    }
    println("Z11 Ошибка: Пользователь $chatId не найден в таблице")
    return false
}



fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, block: Int) {
    println("aaa addScoreForPadezh // Начало добавления балла для пользователя")

    println("a1 Входные параметры: chatId=$chatId, Padezh=$Padezh, filePath=$filePath, block=$block")

    val columnRanges = mapOf(
        1 to mapOf("Именительный" to 1, "Родительный" to 2, "Винительный" to 3, "Дательный" to 4, "Местный" to 5, "Исходный" to 6),
        2 to mapOf("Именительный" to 7, "Родительный" to 8, "Винительный" to 9, "Дательный" to 10, "Местный" to 11, "Исходный" to 12),
        3 to mapOf("Именительный" to 13, "Родительный" to 14, "Винительный" to 15, "Дательный" to 16, "Местный" to 17, "Исходный" to 18)
    )

    val column = columnRanges[block]?.get(Padezh)
    println("a2 Определённая колонка для записи: $column")
    if (column == null) {
        println("❌ Ошибка: колонка для блока $block и падежа $Padezh не найдена.")
        return
    }

    val file = File(filePath)
    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
        println("a3 Проверка наличия листа 'Состояние пользователя'")
        if (sheet == null) {
            println("❌ Ошибка: Лист 'Состояние пользователя' не найден.")
            return
        }
        println("a4 Лист найден. Всего строк: ${sheet.lastRowNum + 1}")

        var userFound = false

        for (row in sheet) {
            val idCell = row.getCell(0)
            println("a5 Проверка ячейки ID в строке ${row.rowNum + 1}")
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
            println("a6 ID в ячейке строки ${row.rowNum + 1}: $idFromCell")

            if (idFromCell == chatId) {
                userFound = true
                println("a7 Пользователь найден в строке ${row.rowNum + 1}. Начинаем обновление.")
                val targetCell = row.getCell(column) ?: row.createCell(column)
                val currentValue = targetCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("a8 Текущее значение в колонке $column: $currentValue")

                targetCell.setCellValue(currentValue + 1)
                println("a9 Новое значение в колонке $column: ${currentValue + 1}")

                safelySaveWorkbook(workbook, filePath)
                println("a10 Изменения сохранены в файле $filePath.")
                return
            }
        }
        if (!userFound) {
            println("⚠️ Пользователь с ID $chatId не найден. Новая запись не создана.")
        }
    }
}
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

// Добавляем инициализацию состояний блоков
fun initializeUserBlockStates(chatId: Long, filePath: String) {
    println("hhh initializeUserBlockStates // Инициализация состояний блоков пользователя")

    println("h1 Входные параметры: chatId=$chatId, filePath=$filePath")

    // Проверяем состояние каждого блока
    val block1Completed = checkUserState(chatId, filePath, block = 1)
    println("h2 Состояние блока 1 для пользователя $chatId: $block1Completed")

    val block2Completed = checkUserState(chatId, filePath, block = 2)
    println("h3 Состояние блока 2 для пользователя $chatId: $block2Completed")

    val block3Completed = checkUserState(chatId, filePath, block = 3)
    println("h4 Состояние блока 3 для пользователя $chatId: $block3Completed")

    println("h5 Предыдущее состояние userBlockCompleted[chatId]: ${userBlockCompleted[chatId]}")

    // Обновляем состояние блоков для пользователя
    userBlockCompleted[chatId] = Triple(block1Completed, block2Completed, block3Completed)
    println("h6 ✅ Обновлённое состояние userBlockCompleted[chatId]: ${userBlockCompleted[chatId]}")
}

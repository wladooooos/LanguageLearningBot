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

val tableFile = "Алгоритм 3.7.xlsx"

fun main() {

    val bot = bot {
        token = "7856005284:AAFVvPnRadWhaotjUZOmFyDFgUHhZ0iGsCo"

        dispatch {
            command("start") {
                val chatId = message.chat.id
                println("3. 🔍 Команда /start от пользователя: $chatId")

                // Полный сброс состояния
                userStates.remove(chatId)
                userPadezh.remove(chatId)
                userWords.remove(chatId)
                userBlocks[chatId] = 1
                userBlockCompleted.remove(chatId)
                userColumnOrder.remove(chatId)

                sendWelcomeMessage(chatId, bot)
                handleBlock(chatId, bot, tableFile, null, null)
            }

            callbackQuery {
                val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery
                val data = callbackQuery.data ?: return@callbackQuery
                println("2. 🔍 Получен callback от пользователя: $chatId, данные: $data")

                println("3. 🔍 Обработка callbackQuery: $data для пользователя $chatId")

                when {
                    data.startsWith("Padezh:") -> {
                        val selectedPadezh = data.removePrefix("Padezh:")
                        println("4. ✅ Выбранный падеж: $selectedPadezh")
                        userPadezh[chatId] = selectedPadezh
                        userStates[chatId] = 0
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Вы выбрали: $selectedPadezh."
                        )
                        if (userWords.containsKey(chatId)) {
                            val (wordUz, wordRus) = userWords[chatId]!!
                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                        } else {
                            bot.sendMessage(
                                chatId = ChatId.fromId(chatId),
                                text = "Теперь выберите слово."
                            )
                            sendWordMessage(chatId, bot, tableFile) // Отправляем клавиатуру со словами
                        }
                    }
                    data.startsWith("word:") -> {
                        println("5. 🔍 Выбор слова: $data")
                        val (wordUz, wordRus) = extractWordsFromCallback(data)
                        println("6. ✅ Выбранное слово: $wordUz ($wordRus)")
                        userWords[chatId] = Pair(wordUz, wordRus)
                        userStates[chatId] = 0
                        sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                    }
                    data.startsWith("next:") -> {
                        println("7. 🔍 Обработка кнопки 'Далее': $data")
                        val params = data.removePrefix("next:").split(":")
                        if (params.size == 2) {
                            val wordUz = params[0]
                            val wordRus = params[1]
                            val currentState = userStates[chatId] ?: 0
                            println("8. 🔎Обновлено текущее состояние пользователя $chatId: ${userStates[chatId]}--------------------------------------------------------------------------------------------")
                            userStates[chatId] = currentState + 1

                            println("🔍 Перед вызовом handleBlock: userBlocks[chatId] = ${userBlocks[chatId]} для пользователя $chatId")
                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                        }
                    }
                    data.startsWith("repeat:") -> {
                        println("9. 🔍 Обработка кнопки 'Повторить': $data")
                        val params = data.removePrefix("repeat:").split(":")
                        if (params.size == 2) {
                            val wordUz = params[0]
                            val wordRus = params[1]
                            userStates[chatId] = 0
                            sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                        }
                    }
                    data == "change_word" -> {
                        println("10. 🔍 Выбор нового слова для пользователя: $chatId")
                        userStates[chatId] = 0
                        sendWordMessage(chatId, bot, tableFile)
                    }
                    data == "change_Padezh" -> {
                        println("11. 🔍 Выбор нового падежа для пользователя: $chatId")
                        userPadezh.remove(chatId)
                        sendPadezhSelection(chatId, bot, tableFile)
                    }
                    data == "reset" -> {
                        println("12. 🔍 Полный сброс данных для пользователя: $chatId")
                        userWords.remove(chatId)
                        userPadezh.remove(chatId)
                        userStates.remove(chatId)
                        sendPadezhSelection(chatId, bot, tableFile)
                    }
                    data == "test" -> {
                        println("13. 🔍 Запрос на тестовую функцию от пользователя: $chatId")
                        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Тест в разработке.")
                    }
                    data == "next_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        val blockStates = userBlockCompleted[chatId] ?: Triple(false, false, false)
                        userWords.remove(chatId)
                        if (currentBlock == 1 && blockStates.first ||
                            currentBlock == 2 && blockStates.second ||
                            currentBlock == 3) {
                            if (currentBlock < 3) {
                                userBlocks[chatId] = currentBlock + 1
                                println("🔍 После переключения блока: userBlocks[chatId] = ${userBlocks[chatId]} для пользователя $chatId")
                                initializeUserBlockStates(chatId, tableFile) // Обновляем состояния
                                sendPadezhSelection(chatId, bot, tableFile)
                            } else {
                                bot.sendMessage(
                                    chatId = ChatId.fromId(chatId),
                                    text = "Вы уже на последнем блоке."
                                )
                            }
                        } else {
                            println("6. ⚠️ Блок $currentBlock не завершен. Сообщаем пользователю.")
                            bot.sendMessage(
                                chatId = ChatId.fromId(chatId),
                                text = "Пройдите все падежи текущего блока, чтобы открыть следующий."
                            )
                        }
                    }

                    data == "prev_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        userWords.remove(chatId)
                        if (currentBlock > 1) {
                            userBlocks[chatId] = currentBlock - 1
                            println("7. 🔍 Возврат на предыдущий блок для пользователя $chatId. Новый блок: ${currentBlock - 1}")
                            sendPadezhSelection(chatId, bot, tableFile)
                        }
                    }
                }
            }
        }
    }

    bot.startPolling()
}
fun handleBlock(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("Вход в функцию handleBlock(). ${userBlocks[chatId]}!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    println("Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")

    val currentBlock = userBlocks[chatId] ?: 1
    println("🔍 Вход в handleBlock: chatId = $chatId, currentBlock = $currentBlock, userPadezh = ${userPadezh[chatId]}")

    initializeUserBlockStates(chatId, filePath)
    println("Состояния блоков пользователя после инициализации: ${userBlockCompleted[chatId]}")

    // Проверяем, выбран ли падеж
    if (userPadezh[chatId] == null) {
        println("⚠️ Падеж не выбран для пользователя $chatId. Отправляем выбор падежа.")
        sendPadezhSelection(chatId, bot, filePath)
        return // Завершаем выполнение, чтобы дождаться выбора падежа
    } else println("AAA падеж выбран ${userPadezh[chatId]}")

    println("ААА Выбираем блок: $currentBlock")
    when (currentBlock) {
        1 -> {
            println("AAA Обрабатываем блок $currentBlock для пользователя $chatId")
            handleBlock1(chatId, bot, filePath, wordUz, wordRus)
        }
        2 -> {
            println("AAA Обрабатываем блок $currentBlock для пользователя $chatId")
            handleBlock2(chatId, bot, filePath, wordUz, wordRus)
        }
        3 -> {
            println("AAA Обрабатываем блок $currentBlock для пользователя $chatId")
            handleBlock3(chatId, bot, filePath, wordUz, wordRus)
        }
        else -> {
            println("Неизвестный блок $currentBlock для пользователя $chatId")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Неизвестный блок: $currentBlock"
            )
        }
    }

    println("Выход из функции handleBlock()")
}


// Функция обработки логики, когда пользователь находится в блоке 1
fun handleBlock1(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    // Аналогичная логика для блока 3
    sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
    // Та же или иная логика для третьего блока
}

fun handleBlock2(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("A. 1. 🔍 Обработка блока 2 для пользователя $chatId")

    // Определяем падеж и соответствующие диапазоны
    val selectedPadezh = userPadezh[chatId]
    println("A. 2. Выбранный падеж: $selectedPadezh")
    if (selectedPadezh == null) {
        println("A. 3. ❌ Ошибка: Падеж не выбран для пользователя $chatId")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
        return
    }

    val blockRanges = mapOf(
        "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7"),
        "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14"),
        "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21"),
        "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28"),
        "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"),
        "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42", "F36-F42")
    )[selectedPadezh] ?: return
    println("A. 4. Диапазоны для выбранного падежа: $blockRanges")

    // Если порядок столбцов для пользователя ещё не создан, создаём и сохраняем его
    if (userColumnOrder[chatId].isNullOrEmpty()) {
        userColumnOrder[chatId] = blockRanges.shuffled().toMutableList()
        println("A. 5. 🔄 Новый перемешанный порядок столбцов для пользователя $chatId: ${userColumnOrder[chatId]}")
    }

    // Получаем текущий индекс (состояние) и список столбцов
    val currentState = userStates[chatId] ?: 0
    val shuffledColumns = userColumnOrder[chatId]!!
    println("A. 6. Текущее состояние (индекс): $currentState")
    println("A. 7. Перемешанный список столбцов: $shuffledColumns")

    // Проверяем, завершены ли все столбцы
    if (currentState >= shuffledColumns.size) {
        println("A. 8. ✅ Все столбцы блока 2 завершены для пользователя $chatId")
        addScoreForPadezh(chatId, selectedPadezh, filePath, block = 2)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    println("A. 9. 🔁 Начинаем обработку столбцов от $currentState до ${shuffledColumns.size - 1}")
    for (i in currentState until shuffledColumns.size) {
        val range = shuffledColumns[i]
        println("A. 10. Текущий диапазон для обработки: $range")

        val messageText = generateMessageFromRange(filePath, "Существительные 2", range, wordUz, wordRus)
        println("A. 11. Сформированное сообщение для диапазона $range: $messageText")

        // Проверяем, последний ли это столбец
        val isLastMessage = i == shuffledColumns.size - 1
        println("A. 12. Является ли это последний столбец: $isLastMessage")

        // Отправляем сообщение
        if (isLastMessage) {
            println("A. 13. 📨 Отправка последнего сообщения без кнопки для пользователя $chatId")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = messageText,
                parseMode = ParseMode.MARKDOWN_V2
            )
            println("A. 14. ✅ Последний столбец обработан: $range")
            addScoreForPadezh(chatId, selectedPadezh, filePath, block = 2)
            sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        } else {
            println("A. 15. 📨 Отправка сообщения с кнопкой 'Далее' для диапазона $range")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = messageText,
                parseMode = ParseMode.MARKDOWN_V2,
                replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                    InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus")
                )
            )
            println("A. 16. ✅ Сообщение для диапазона $range отправлено с кнопкой")
        }

        // Останавливаем цикл после отправки первого сообщения, чтобы дождаться действия пользователя
        if (!isLastMessage) {
            println("A. 18. Цикл завершен на столбце $range, ждем действия пользователя")
            return
        }
    }
    println("A. 19. 🔄 Завершение обработки блока 2 для пользователя $chatId")
}




// Функция обработки логики, когда пользователь находится в блоке 3
fun handleBlock3(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    // Аналогичная логика для блока 3
    sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
    // Та же или иная логика для третьего блока
}

// Формирует и отправляет клавиатуру для выбора падежей текущего блока.
fun sendPadezhSelection(chatId: Long, bot: Bot, filePath: String) {
    println("🔍 Формируем клавиатуру выбора падежа для пользователя $chatId")

    // Сбрасываем текущий выбранный падеж
    userPadezh.remove(chatId)
    userStates.remove(chatId)

    val currentBlock = userBlocks[chatId] ?: 1 // Получаем текущий блок пользователя
    println("🔍 Текущий блок: $currentBlock")

    val PadezhColumns = getPadezhColumnsForBlock(currentBlock)
    if (PadezhColumns == null) {
        println("⚠️ Ошибка: блок $currentBlock не найден.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: невозможно загрузить данные блока.")
        return
    }

    val userScores = getUserScoresForBlock(chatId, filePath, PadezhColumns)
    val buttons = generatePadezhSelectionButtons(currentBlock, PadezhColumns, userScores)

    // Отправляем сообщение
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите падеж для изучения блока $currentBlock:",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
}


// Функция для получения колонок текущего блока
fun getPadezhColumnsForBlock(block: Int): Map<String, Int>? {
    val columnRanges = mapOf(
        1 to mapOf("Именительный" to 1, "Родительный" to 2, "Винительный" to 3, "Дательный" to 4, "Местный" to 5, "Исходный" to 6),
        2 to mapOf("Именительный" to 7, "Родительный" to 8, "Винительный" to 9, "Дательный" to 10, "Местный" to 11, "Исходный" to 12),
        3 to mapOf("Именительный" to 13, "Родительный" to 14, "Винительный" to 15, "Дательный" to 16, "Местный" to 17, "Исходный" to 18)
    )
    return columnRanges[block]
}

// Функция для чтения баллов пользователя для блока
fun getUserScoresForBlock(chatId: Long, filePath: String, PadezhColumns: Map<String, Int>): Map<String, Int> {
    println("🔍 Считываем баллы пользователя $chatId из файла $filePath")
    val file = File(filePath)
    if (!file.exists()) {
        println("❌ Файл $filePath не найден.")
        return emptyMap()
    }

    val scores = mutableMapOf<String, Int>()

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
        if (sheet == null) {
            println("❌ Лист 'Состояние пользователя' не найден.")
            return emptyMap()
        }

        // Ищем строку пользователя
        for (row in sheet) {
            val idCell = row.getCell(0)
            val chatIdXlsx = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                else -> null
            }

            if (chatId == chatIdXlsx) {
                println("✅ Найдена строка пользователя $chatIdXlsx. Извлекаем баллы...")
                for ((PadezhName, colIndex) in PadezhColumns) {
                    val cell = row.getCell(colIndex)
                    val score = cell?.numericCellValue?.toInt() ?: 0
                    scores[PadezhName] = score
                }
                break
            }
        }
    }

    println("✅ Считанные баллы для пользователя $chatId: $scores")
    return scores
}

// Функция для формирования кнопок выбора падежей
fun generatePadezhSelectionButtons(
    currentBlock: Int,
    PadezhColumns: Map<String, Int>,
    userScores: Map<String, Int>
): List<List<InlineKeyboardButton>> {
    println("🔧 Генерируем кнопки для выбора падежей в блоке $currentBlock")
    val buttons = PadezhColumns.keys.map { PadezhName ->
        val score = userScores[PadezhName] ?: 0 // Если баллов нет, используем 0
        InlineKeyboardButton.CallbackData("$PadezhName [$score]", "Padezh:$PadezhName")
    }.map { listOf(it) }.toMutableList()

    // Добавляем кнопки для переключения блоков
    if (currentBlock > 1) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
    }
    if (currentBlock < 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
    }

    println("✅ Кнопки успешно сгенерированы")
    return buttons
}

// sendWelcomeMessage: Отправляет приветственное сообщение с кнопкой /start.
fun sendWelcomeMessage(chatId: Long, bot: Bot) {
    println("22. 🔔 Отправка приветственного сообщения пользователю $chatId.")
    val keyboardMarkup = KeyboardReplyMarkup(
        keyboard = generateUsersButton(),
        resizeKeyboard = true
    )
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = """Здравствуйте!
Я бот, помогающий изучать узбекский язык!""",
        replyMarkup = keyboardMarkup
    )
}

// generateUsersButton: Генерирует кнопку /start для быстрого доступа.
fun generateUsersButton(): List<List<KeyboardButton>> {
    println("23. 🛠️ Генерация кнопки /start.")
    return listOf(
        listOf(KeyboardButton("/start"))
    )
}

// sendWordMessage: Отправляет клавиатуру с выбором слов.
fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
    println("24. 🔍 Отправка клавиатуры с выбором слов для пользователя $chatId.")
    if (!File(filePath).exists()) {
        println("25. ❌ Ошибка: файл $filePath не найден.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка: файл с данными не найден."
        )
        return
    }

    val inlineKeyboard = try {
        println("26. 🛠️ Генерация клавиатуры из файла $filePath.")
        createWordSelectionKeyboardFromExcel(filePath, "Существительные")
    } catch (e: Exception) {
        println("27. ❌ Ошибка при обработке данных из файла: ${e.message}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка при обработке данных: ${e.message}"
        )
        return
    }

    println("28. 📨 Отправка сообщения с клавиатурой.")
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите слово из списка:",
        replyMarkup = inlineKeyboard
    )
}


// createWordSelectionKeyboardFromExcel: Создает клавиатуру из данных в Excel-файле.
fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
    println("29. 🔍 Начало создания клавиатуры из файла: $filePath, лист: $sheetName")

    // Проверка существования файла
    println("30. 🔎 Проверяем существование файла $filePath")
    val file = File(filePath)
    if (!file.exists()) {
        println("31. ❌ Файл $filePath не найден")
        throw IllegalArgumentException("Файл $filePath не найден")
    }
    println("32. ✅ Файл найден: $filePath")

    // Открытие файла и получение листа
    println("33. 📂 Открываем файл Excel")
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName)
        ?: throw IllegalArgumentException("34. ❌ Лист $sheetName не найден")
    println("35. ✅ Лист найден: $sheetName")

    println("36. 📜 Всего строк в листе: ${sheet.lastRowNum + 1}")

    // Генерация случайных строк
    println("37. 🎲 Генерация случайных строк для кнопок")
    val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
    println("38. 🎲 Выбраны строки: $randomRows")

    // Создание кнопок
    println("39. 🔨 Начинаем обработку строк для создания кнопок")
    val buttons = randomRows.mapNotNull { rowIndex ->
        val row = sheet.getRow(rowIndex)
        if (row == null) {
            println("40. ⚠️ Строка $rowIndex отсутствует, пропускаем")
            return@mapNotNull null
        }

        val wordUz = row.getCell(0)?.toString()?.trim()
        val wordRus = row.getCell(1)?.toString()?.trim()

        if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
            println("41. ⚠️ Некорректные данные в строке $rowIndex: wordUz = $wordUz, wordRus = $wordRus. Пропускаем")
            return@mapNotNull null
        }

        println("42. ✅ Обработана строка $rowIndex: wordUz = $wordUz, wordRus = $wordRus")
        InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
    }.chunked(2) // Группируем кнопки по 2 в строке

    println("43. ✅ Кнопки успешно созданы. Количество строк кнопок: ${buttons.size}")

    // Закрытие файла
    println("44. 📕 Закрываем файл Excel")
    workbook.close()
    println("45. ✅ Файл Excel успешно закрыт")

    // Генерация клавиатуры завершена
    println("46. 🔑 Генерация клавиатуры завершена")
    return InlineKeyboardMarkup.create(buttons)
}

// extractWordsFromCallback: Извлекает узбекское и русское слово из callback data.
fun extractWordsFromCallback(data: String): Pair<String, String> {
    println("47. 🔍 Извлечение слов из callback data: $data")
    val content = data.substringAfter("word:(").substringBefore(")")
    val wordUz = content.substringBefore(" - ").trim()
    val wordRus = content.substringAfter(" - ").trim()
    println("48. ✅ Извлечённые слова: wordUz = $wordUz, wordRus = $wordRus")
    return wordUz to wordRus
}

// sendStateMessage: Отправляет сообщение по текущему состоянию и падежу.
fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("49. 🔍 Формируем сообщение для пользователя: $chatId, слово: $wordUz, перевод: $wordRus")

    val (selectedPadezh, rangesForPadezh, currentState) = validateUserState(chatId, bot) ?: return
    processStateAndSendMessage(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState)
}

fun validateUserState(chatId: Long, bot: Bot): Triple<String, List<String>, Int>? {
    val selectedPadezh = userPadezh[chatId]
    if (selectedPadezh == null) {
        println("50. ❌ Ошибка: выбранный падеж отсутствует.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
        return null
    }
    println("51. ✅ Выбранный падеж: $selectedPadezh")

    val rangesForPadezh = PadezhRanges[selectedPadezh]
    if (rangesForPadezh == null) {
        println("52. ❌ Ошибка: диапазоны для падежа $selectedPadezh отсутствуют.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: диапазоны для падежа не найдены.")
        return null
    }

    val currentState = userStates[chatId] ?: 0
    println("53. 🔎 Текущее состояние: $currentState")

    return Triple(selectedPadezh, rangesForPadezh, currentState)
}

fun processStateAndSendMessage(
    chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
    selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int
) {
    if (currentState >= rangesForPadezh.size) {
        println("54. ✅ Все этапы завершены для падежа: $selectedPadezh")
        addScoreForPadezh(chatId, selectedPadezh, filePath)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    val range = rangesForPadezh[currentState]
    println("55. 🔍 Генерация сообщения для диапазона: $range")
    val currentBlock = userBlocks[chatId] ?: 1 // Получаем текущий блок пользователя (по умолчанию 1)
    val listName = when (currentBlock) {
        1 -> "Существительные 1"
        2 -> "Существительные 2"
        3 -> "Существительные 3"
        else -> "Существительные 1" // Значение по умолчанию, если блок не указан
    }

    val messageText = try {
        generateMessageFromRange(filePath, listName, range, wordUz, wordRus)
    } catch (e: Exception) {
        println("56. ❌ Ошибка при генерации сообщения: ${e.message}")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        return
    }

    println("57. ✅ Сообщение сгенерировано.")
    sendMessageOrNextStep(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState, messageText)
}

fun sendMessageOrNextStep(
    chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?,
    selectedPadezh: String, rangesForPadezh: List<String>, currentState: Int, messageText: String
) {
    if (currentState == rangesForPadezh.size - 1) {
        println("571 $messageText")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2
        )
        println("58. ✅ Последний этап завершен. Добавляем балл и отправляем финальное меню.")
        val currentBlock = userBlocks[chatId] ?: 1
        println("59. ✅ Уточняем текущий блок пользователя.")
        addScoreForPadezh(chatId, selectedPadezh, filePath, currentBlock)
        println("60. ✅ Добавляем балл для текущего блока.")
        println("61. ✅ Отправка финального меню...")
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        println("62. ✅ Финальное меню отправлено.")
    } else {
        println("572 $messageText")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2,
            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus")
            )
        )
    }
}



// sendFinalButtons: Отправляет клавиатуру с финальными вариантами действий.
fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
    println("92. 🔍 Формируем меню финальных действий для пользователя $chatId")
    val showNextStep = checkUserState(chatId, filePath) // Проверяем состояние пользователя
    println("93. 🔎 Состояние пользователя. Добавить кнопку 'Следующий блок': $showNextStep")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Повторить", "repeat:$wordUz:$wordRus")),
        listOf(InlineKeyboardButton.CallbackData("Изменить слово", "change_word")),
        listOf(InlineKeyboardButton.CallbackData("Изменить падеж", "change_Padezh")),
        //listOf(InlineKeyboardButton.CallbackData("Изменить падеж и слово", "reset"))
    )
    val currentBlock = userBlocks[chatId] ?: 1
    if (currentBlock > 1) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
    }
    if (currentBlock < 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
    }

    println("95. 📨 Отправляем сообщение с финальным меню.")
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
}

fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String?, wordRus: String?): String {
    println("96. 🔍 Генерация сообщения. Файл: $filePath, Лист: $sheetName, Диапазон: $range, Слова: $wordUz, $wordRus.")
    val file = File(filePath)
    if (!file.exists()) {
        println("97. ❌ Файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден")
    }

    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName)
    if (sheet == null) {
        println("98. ❌ Лист $sheetName не найден.")
        throw IllegalArgumentException("Лист $sheetName не найден")
    }
    println("99. ✅ Лист $sheetName найден. Начинаем извлечение ячеек.")

    val cells = extractCellsFromRange(sheet, range, wordUz)
    println("100. 📜 Извлечённые ячейки: $cells")

    val firstCell = cells.firstOrNull() ?: "" // Первая ячейка, не блюрим
    println("101. 🔑 Первый элемент: \"$firstCell\"")

    val messageBody = cells.drop(1).joinToString("\n") // Остальные элементы объединяются
    println("102. 📄 Содержимое тела сообщения:\n$messageBody")

    workbook.close()
    println("103. 📕 Файл закрыт. Генерация завершена.")

    return listOf(firstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n") // Объединяем первую ячейку и тело
}


fun String.escapeMarkdownV2(): String {
    println("🔧 Экранирование Markdown для строки: \"$this\"")
    return this.replace("\\", "\\\\")
        .replace("_", "\\_")
        .replace("*", "\\*")
        .replace("[", "\\[")
        .replace("]", "\\]")
        .replace("(", "\\(") // Экранирование открывающей скобки
        .replace(")", "\\)") // Экранирование закрывающей скобки
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
}


fun adjustWordUz(content: String, wordUz: String?): String {
    println("105. 🔍 Обработка слова \"$wordUz\" в контексте \"$content\"")

    // Вспомогательная функция для проверки, является ли символ гласным
    fun Char.isVowel() = this.lowercaseChar() in "aeiouаеёиоуыэюя"

    // Строим результат через StringBuilder
    return buildString {
        var i = 0
        while (i < content.length) {
            val char = content[i]

            when {
                // Если встречается '+', применяем правила замены
                char == '+' && i + 1 < content.length -> {
                    val nextChar = content[i + 1]
                    val lastChar = wordUz?.lastOrNull()

                    val replacement = when {
                        lastChar != null && lastChar.isVowel() && nextChar.isVowel() -> "s"
                        lastChar != null && !lastChar.isVowel() && !nextChar.isVowel() -> "i"
                        else -> ""
                    }

                    append(replacement)
                    append(nextChar) // Добавляем символ после `+` (например, `m`).
                    i++ // Пропускаем символ `nextChar` для корректной обработки.
                }

                // Если встречается '*', заменяем на wordUz
                char == '*' -> {
                    append(wordUz)
                }

                // В остальных случаях просто добавляем символ
                else -> {
                    append(char)
                }
            }

            i++
        }
    }
}





// extractCellsFromRange: Извлекает и обрабатывает ячейки из указанного диапазона листа Excel.
fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String?): List<String> {
    println("108. 🔍 Извлечение ячеек из диапазона $range для слова \"$wordUz\"")
    val (start, end) = range.split("-").map { it.replace(Regex("[A-Z]"), "").toInt() - 1 }
    val column = range[0] - 'A'
    println("109. 📌 Диапазон строк: $start-$end, Колонка: $column")

    return (start..end).mapNotNull { rowIndex ->
        val row = sheet.getRow(rowIndex)
        if (row == null) {
            println("110. ⚠️ Строка $rowIndex отсутствует. Пропускаем.")
            return@mapNotNull null
        }
        val cell = row.getCell(column)
        if (cell == null) {
            println("111. ⚠️ Ячейка в строке $rowIndex отсутствует. Пропускаем.")
            return@mapNotNull null
        }
        val processed = processCellContent(cell, wordUz)
        println("112. ✅ Обработанная ячейка в строке $rowIndex: \"$processed\"")
        processed
    }
}

fun checkUserState(chatId: Long, filePath: String, sheetName: String = "Состояние пользователя" ): Boolean {
    println("113. 🔍 Проверяем состояние пользователя. ID: $chatId, Файл: $filePath, Лист: $sheetName")

    val file = File(filePath)
    var userRow: Row? = null
    var allCompleted = false

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
            ?: throw IllegalArgumentException("114. ❌ Лист $sheetName не найден")
        println("115. ✅ Лист найден: $sheetName")

        println("116. 🔍 Начинаем поиск пользователя по ID или пустой строки")
        for (rowIndex in 1..sheet.lastRowNum.coerceAtLeast(1)) {
            val row = sheet.getRow(rowIndex) ?: sheet.createRow(rowIndex)
            val idCell = row.getCell(0) ?: row.createCell(0)

            val currentId = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toLongOrNull()
                else -> null
            }
            println("117. 🔎 Проверяем строку: ${rowIndex + 1}, ID: $currentId")

            if (currentId == chatId) {
                println("118. ✅ Пользователь найден в строке: ${rowIndex + 1}")
                userRow = row
                break
            }

            if (currentId == null || currentId == 0L) {
                println("119. ⚠️ Пустая ячейка. Записываем нового пользователя.")
                idCell.setCellValue(chatId.toDouble())
                for (i in 1..6) {
                    row.createCell(i).setCellValue(0.0)
                }
                safelySaveWorkbook(workbook, filePath)
                println("120. ✅ Пользователь добавлен в строку: ${rowIndex + 1}")
                return false
            }
        }

        if (userRow != null) {
            println("121. 🔍 Проверяем значения в колонках Им., Род., Вин., Дат., Мест., Исх.")
            allCompleted = (1..6).all { index ->
                val cell = userRow!!.getCell(index)
                val value = cell?.toString()?.toDoubleOrNull() ?: 0.0
                println("122. 🔎 Колонка ${index + 1}: значение = $value")
                value > 0
            }
        } else {
            val newRowIndex = sheet.physicalNumberOfRows
            val newRow = sheet.createRow(newRowIndex)
            newRow.createCell(0).setCellValue(chatId.toDouble())
            for (i in 1..6) {
                newRow.createCell(i).setCellValue(0.0)
            }
            safelySaveWorkbook(workbook, filePath)
            println("123. ✅ Пользователь добавлен в строку: ${newRowIndex + 1}")
            return false
        }
    }
    println("$userBlockCompleted---------------------------------------------------------------------------------------------------------------------------")
    initializeUserBlockStates(chatId, tableFile) // Обновляем состояния
    println("$userBlockCompleted---------------------------------------------------------------------------------------------------------------------------")
    println("124. ✅ Проверка завершена. Все шаги завершены: $allCompleted")
    return allCompleted
}

fun safelySaveWorkbook(workbook: org.apache.poi.ss.usermodel.Workbook, filePath: String) {
    val tempFile = File("${filePath}.tmp")
    println("125. 📂 Начало сохранения файла. Временный файл: ${tempFile.path}")

    try {
        FileOutputStream(tempFile).use { outputStream ->
            println("126. 💾 Пишем данные в временный файл: ${tempFile.path}")
            workbook.write(outputStream)
        }
        println("127. ✅ Данные успешно записаны во временный файл.")

        val originalFile = File(filePath)
        println("128. 📂 Проверяем наличие оригинального файла: ${originalFile.path}")
        if (originalFile.exists()) {
            println("129. 🗑️ Оригинальный файл существует. Удаляем: ${originalFile.path}")
            if (!originalFile.delete()) {
                println("130. ❌ Не удалось удалить оригинальный файл: ${originalFile.path}")
                throw IllegalStateException("Не удалось удалить старый файл: $filePath")
            }
            println("131. ✅ Оригинальный файл успешно удалён.")
        }

        println("132. 🔄 Переименование временного файла в оригинальный: ${tempFile.path} -> ${originalFile.path}")
        if (!tempFile.renameTo(originalFile)) {
            println("133. ❌ Не удалось переименовать временный файл: ${tempFile.path}")
            throw IllegalStateException("Не удалось переименовать временный файл: ${tempFile.path}")
        }
        println("134. ✅ Файл успешно сохранён: $filePath")
    } catch (e: Exception) {
        println("135. ❌ Ошибка при сохранении файла: ${e.message}")
        println("136. 🗑️ Удаляем временный файл: ${tempFile.path}")
        tempFile.delete()
        throw e
    }
}



fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, sheetName: String = "Состояние пользователя") {
    println("163. 🔍 Начинаем добавление балла. Пользователь: $chatId, Падеж: $Padezh")

    val file = File(filePath)
    if (!file.exists()) {
        println("164. ❌ Файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден.")
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
            ?: throw IllegalArgumentException("165. ❌ Лист $sheetName не найден")
        println("166. ✅ Лист найден: $sheetName")

        val PadezhColumnIndex = when (Padezh) {
            "Именительный" -> 1
            "Родительный" -> 2
            "Винительный" -> 3
            "Дательный" -> 4
            "Местный" -> 5
            "Исходный" -> 6
            else -> throw IllegalArgumentException("167. ❌ Неизвестный падеж: $Padezh")
        }

        println("168. 🔎 Ищем пользователя $chatId в таблице...")
        for (rowIndex in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue
            val idCell = row.getCell(0)
            val currentId = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            println("169. 🔎 Проверяем строку ${rowIndex + 1}: ID = $currentId.")
            if (currentId == chatId) {
                println("170. ✅ Пользователь найден в строке ${rowIndex + 1}")
                val PadezhCell = row.getCell(PadezhColumnIndex) ?: row.createCell(PadezhColumnIndex)
                val currentScore = PadezhCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("171. 🔎 Текущее значение: $currentScore")
                PadezhCell.setCellValue(currentScore + 1)
                println("172. ✅ Новое значение: ${currentScore + 1}")
                safelySaveWorkbook(workbook, filePath)
                println("173. ✅ Балл добавлен и изменения сохранены.")
                return
            }
        }
        println("174. ⚠️ Пользователь $chatId не найден. Этого не должно происходить на этом этапе.")
    }
}

fun checkUserState(chatId: Long, filePath: String, sheetName: String = "Состояние пользователя", block: Int = 1): Boolean {
    //println("175. 🔍 Проверка состояния пользователя. ID: $chatId, файл: $filePath, лист: $sheetName, блок: $block.")

    val columnRanges = mapOf(
        1 to (1..6),
        2 to (7..12),
        3 to (13..18)
    )
    val columns = columnRanges[block] ?: run {
        //println("176. ⚠️ Неизвестный блок: $block. Возвращаем false.")
        return false
    }

    //println("177. ✅ Диапазон колонок для блока $block: $columns.")

    val file = File(filePath)
    if (!file.exists()) {
        //println("178. ❌ Файл $filePath не найден. Возвращаем false.")
        return false
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            //println("179. ❌ Лист $sheetName не найден. Возвращаем false.")
            return false
        }

        //println("180. ✅ Лист $sheetName найден. Всего строк: ${sheet.lastRowNum + 1}.")

        for (row in sheet) {
            val idCell = row.getCell(0) ?: continue
            val chatIdFromCell = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            //println("181. 🔎 Строка ${row.rowNum + 1}: ID в ячейке = $chatIdFromCell.")
            if (chatIdFromCell == chatId) {
                //println("182. ✅ Пользователь найден в строке ${row.rowNum + 1}. Проверяем выполнение всех условий...")

                val allColumnsCompleted = columns.all { colIndex ->
                    val cell = row.getCell(colIndex)
                    val value = cell?.numericCellValue ?: 0.0
                    //println("183. 📊 Колонка $colIndex: значение = $value. Выполнено? ${value > 0}")
                    value > 0
                }

                println("184. ✅ Результат проверки для пользователя $chatId: $allColumnsCompleted. ЗАВЕРШАЕМ ПРОВЕРКУ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                return allColumnsCompleted
            }
        }
    }

    println("185. ⚠️ Пользователь $chatId не найден в таблице. Возвращаем false.")
    return false
}


fun addScoreForPadezh(chatId: Long, Padezh: String, filePath: String, block: Int) {
    println("🔍 Начинаем добавление балла для пользователя $chatId. Падеж: $Padezh, Файл: $filePath, Блок: $block.")

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
    println("✅ Определена колонка для записи: $column.")

    val file = File(filePath)
    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
        if (sheet == null) {
            println("❌ Ошибка: Лист 'Состояние пользователя' не найден.")
            return
        }
        println("✅ Лист 'Состояние пользователя' найден. Всего строк в листе: ${sheet.lastRowNum + 1}.")

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
                println("❌ Ошибка: не удалось прочитать ID из ячейки в строке ${row.rowNum + 1}. Значение: ${idCell}.")
                continue
            }

            println("🔎 Проверяем строку ${row.rowNum + 1}. ID в ячейке: $idFromCell.")

            if (idFromCell == chatId) {
                userFound = true
                println("✅ Пользователь найден в строке ${row.rowNum + 1}. Начинаем обновление.")
                val targetCell = row.getCell(column) ?: row.createCell(column)
                val currentValue = targetCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("🔎 Текущее значение в колонке $column: $currentValue.")

                targetCell.setCellValue(currentValue + 1)
                println("✅ Новое значение в колонке $column: ${currentValue + 1}. Сохраняем файл.")

                safelySaveWorkbook(workbook, filePath)
                println("✅ Изменения сохранены в файле $filePath.")
                return
            }
        }
        if (!userFound) {
            println("⚠️ Пользователь с ID $chatId не найден. Новая запись не создана.")
        }
    }
}


fun processCellContent(cell: Cell?, wordUz: String?): String {
    println("206. 🔍 Обработка ячейки: $cell")
    if (cell == null) {
        println("207. ⚠️ Ячейка пуста. Возвращаем пустую строку.")
        return ""
    }

    val richText = cell.richStringCellValue as XSSFRichTextString
    val text = richText.string
    println("208. 📜 Содержимое ячейки: \"$text\".")

    val runs = richText.numFormattingRuns()
    println("209. 📊 Количество форматированных участков: $runs")

    return if (runs == 0) {
        processCellWithoutRuns(cell, text, wordUz)
    } else {
        processFormattedRuns(richText, text, wordUz)
    }
}

// Обработка ячейки без форматированных участков
fun processCellWithoutRuns(cell: Cell, text: String, wordUz: String?): String {
    val font = getCellFont(cell)
    val isRed = font != null && getFontColor(font) == "#FF0000"

    return if (isRed) {
        println("🔴 Вся ячейка имеет красный цвет. Блюрим текст.")
        "||${adjustWordUz(text, wordUz).escapeMarkdownV2()}||"
    } else {
        println("⚪ Текст не красный. Обрабатываем как есть.")
        adjustWordUz(text, wordUz).escapeMarkdownV2()
    }
}

// Обработка форматированных участков текста
fun processFormattedRuns(richText: XSSFRichTextString, text: String, wordUz: String?): String {
    println("210. 🔧 Обработка форматированных участков текста.")

    return buildString {
        for (i in 0 until richText.numFormattingRuns()) {
            val start = richText.getIndexOfFormattingRun(i)
            val end = if (i + 1 < richText.numFormattingRuns()) richText.getIndexOfFormattingRun(i + 1) else text.length
            val substring = text.substring(start, end)

            val font = richText.getFontOfFormattingRun(i) as? XSSFFont
            val colorHex = font?.let { getFontColor(it) } ?: "Цвет не определён"
            println("    🎨 Цвет участка $i: $colorHex")

            val adjustedSubstring = adjustWordUz(substring, wordUz)

            if (colorHex == "#FF0000") {
                println("    🔴 Текст участка \"$substring\" красный. Добавляем блюр.")
                append("||${adjustedSubstring.escapeMarkdownV2()}||")
            } else {
                append(adjustedSubstring.escapeMarkdownV2())
            }
        }
    }.also {
        println("213. ✅ Результат обработки форматированных участков: \"$it\".")
    }
}

// Получение шрифта ячейки
fun getCellFont(cell: Cell): XSSFFont? {
    val workbook = cell.sheet.workbook as? org.apache.poi.xssf.usermodel.XSSFWorkbook ?: return null
    val fontIndex = cell.cellStyle.fontIndexAsInt
    return workbook.getFontAt(fontIndex) as? XSSFFont
}






// Функция для извлечения цвета шрифта
fun getFontColor(font: XSSFFont): String {
    println("🔍 Извлекаем цвет шрифта...")

    val xssfColor = font.xssfColor // Получаем цвет шрифта
    if (xssfColor == null) {
        println("⚠️ Цвет шрифта не определён.")
        return "Цвет не определён"
    }

    val rgb = xssfColor.rgb // Проверяем наличие RGB-цвета
    return if (rgb != null) {
        val colorHex = rgb.joinToString(prefix = "#", separator = "") { "%02X".format(it) }
        println("🎨 Цвет шрифта: $colorHex")
        colorHex
    } else {
        println("⚠️ RGB не найден.")
        "Цвет не определён"
    }
}

// Вспомогательный метод для обработки цветов с учётом оттенков
fun XSSFColor.getRgbWithTint(): ByteArray? {
    return rgb?.let { baseRgb ->
        val tint = this.tint
        if (tint != 0.0) {
            baseRgb.map { (it and 0xFF.toByte()) * (1 + tint) }.map { it.coerceIn(0.0, 255.0).toInt().toByte() }.toByteArray()
        } else {
            baseRgb
        }
    }
}

// Добавляем инициализацию состояний блоков
fun initializeUserBlockStates(chatId: Long, filePath: String) {
    println("1. 🔍 Инициализация состояний блоков для пользователя $chatId")

    // Проверяем состояние каждого блока
    val block1Completed = checkUserState(chatId, filePath, block = 1)
    val block2Completed = checkUserState(chatId, filePath, block = 2)
    val block3Completed = checkUserState(chatId, filePath, block = 3)

    println("🔍 Завершение блоков: Block1 = $block1Completed, Block2 = $block2Completed, Block3 = $block3Completed для пользователя $chatId")
    println("🔍 Предыдущее состояние userBlockCompleted[chatId]: ${userBlockCompleted[chatId]}")

    // Обновляем состояние блоков для пользователя
    userBlockCompleted[chatId] = Triple(block1Completed, block2Completed, block3Completed)

    println("2. ✅ Обновленное состояние userBlockCompleted[chatId]: ${userBlockCompleted[chatId]}")
}

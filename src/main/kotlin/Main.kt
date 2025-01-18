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


val caseRanges = mapOf(
    "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7"),
    "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14"),
    "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21"),
    "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28"),
    "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35"),
    "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42")
)

val userStates = mutableMapOf<Long, Int>()
val userCases = mutableMapOf<Long, String>() // Хранение выбранного падежа для каждого пользователя
val userWords = mutableMapOf<Long, Pair<String, String>>() // Хранение выбранного слова для каждого пользователя
val userBlocks = mutableMapOf<Long, Int>() // Хранение текущего блока для каждого пользователя

val tableFile = "Алгоритм 3.6.xlsx"

fun main() {

    val bot = bot {
        token = "7856005284:AAFVvPnRadWhaotjUZOmFyDFgUHhZ0iGsCo"

        dispatch {
            command("start") {
                val userId = message.chat.id
                println("1. 🔍 Команда /start от пользователя: $userId")
                userStates[userId] = 0
                userWords.remove(userId) // Удаляем старые данные о слове
                sendWelcomeMessage(userId, bot)
                sendCaseSelection(userId, bot, tableFile) // Добавляем tableFile как filePath
            }

            callbackQuery {
                val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery
                val data = callbackQuery.data ?: return@callbackQuery
                println("2. 🔍 Получен callback от пользователя: $chatId, данные: $data")

                println("3. 🔍 Обработка callbackQuery: $data для пользователя $chatId")

                when {
                    data.startsWith("case:") -> {
                        val selectedCase = data.removePrefix("case:")
                        println("4. ✅ Выбранный падеж: $selectedCase")
                        userCases[chatId] = selectedCase
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Вы выбрали: $selectedCase."
                        )
                        if (userWords.containsKey(chatId)) {
                            val (wordUz, wordRus) = userWords[chatId]!!
                            userStates[chatId] = 0
                            bot.sendMessage(
                                chatId = ChatId.fromId(chatId),
                                text = "Начнем с уже выбранного слова: $wordUz ($wordRus)."
                            )
                            sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                        } else {
                            bot.sendMessage(
                                chatId = ChatId.fromId(chatId),
                                text = "Теперь выберите слово."
                            )
                            sendWordMessage(chatId, bot, tableFile)
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
                            println("8. 🔎 Текущее состояние: $currentState. Обновляем.")
                            userStates[chatId] = currentState + 1
                            sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
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
                    data == "change_case" -> {
                        println("11. 🔍 Выбор нового падежа для пользователя: $chatId")
                        sendCaseSelection(chatId, bot, tableFile)
                    }
                    data == "reset" -> {
                        println("12. 🔍 Полный сброс данных для пользователя: $chatId")
                        userWords.remove(chatId)
                        userCases.remove(chatId)
                        userStates.remove(chatId)
                        sendCaseSelection(chatId, bot, tableFile)
                    }
                    data == "test" -> {
                        println("13. 🔍 Запрос на тестовую функцию от пользователя: $chatId")
                        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Тест в разработке.")
                    }
                    data == "next_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        if (currentBlock < 3) {
                            userBlocks[chatId] = currentBlock + 1
                            println("14. 🔍 Переход на следующий блок для пользователя $chatId. Новый блок: ${currentBlock + 1}")
                            sendCaseSelection(chatId, bot, tableFile)
                        } else {
                            val totalScore = calculateTotalScore(chatId, tableFile)
                            bot.sendMessage(
                                chatId = ChatId.fromId(chatId),
                                text = "Вы прошли все блоки. Сейчас у вас $totalScore баллов. Для прохождения теста наберите 50 баллов."
                            )
                        }
                    }

                    data == "prev_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        if (currentBlock > 1) {
                            userBlocks[chatId] = currentBlock - 1
                            println("15. 🔍 Возврат на предыдущий блок для пользователя $chatId. Новый блок: ${currentBlock - 1}")
                            sendCaseSelection(chatId, bot, tableFile)
                        }
                    }
                    else -> {
                        println("16. ⚠️ Неизвестный callbackQuery: $data")
                    }
                }
            }
        }
    }

    bot.startPolling()
}
// sendCaseSelection: Отправляет пользователю клавиатуру для выбора падежа.
fun sendCaseSelection(chatId: Long, bot: Bot, filePath: String) {
    val currentBlock = userBlocks[chatId] ?: 1
    println("17. 🔍 Пользователь $chatId выбирает падеж. Текущий блок: $currentBlock")

    val showNextStep = checkUserState(chatId, filePath, block = currentBlock)
    println("18. 🔎 Доступен ли следующий шаг? $showNextStep")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Именительный падеж", "case:Именительный")),
        listOf(InlineKeyboardButton.CallbackData("Родительный падеж", "case:Родительный")),
        listOf(InlineKeyboardButton.CallbackData("Винительный падеж", "case:Винительный")),
        listOf(InlineKeyboardButton.CallbackData("Дательный падеж", "case:Дательный")),
        listOf(InlineKeyboardButton.CallbackData("Местный падеж", "case:Местный")),
        listOf(InlineKeyboardButton.CallbackData("Исходный падеж", "case:Исходный"))
    )

    if (currentBlock > 1) {
        println("19. 🔙 Добавляем кнопку 'Предыдущий блок'.")
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Предыдущий блок", "prev_block")))
    }
    if (showNextStep) {
        println("20. 🔜 Добавляем кнопку 'Следующий блок'.")
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Следующий блок", "next_block")))
    }

    println("21. 📨 Отправляем сообщение с выбором падежа.")
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите падеж для изучения блока $currentBlock:",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
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
fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String, wordRus: String) {
    println("49. 🔍 Формируем сообщение для пользователя: $chatId, слово: $wordUz, перевод: $wordRus")

    val selectedCase = userCases[chatId]
    if (selectedCase == null) {
        println("50. ❌ Ошибка: выбранный падеж отсутствует.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
        return
    }

    println("51. ✅ Выбранный падеж: $selectedCase")

    val rangesForCase = caseRanges[selectedCase]
    if (rangesForCase == null) {
        println("52. ❌ Ошибка: диапазоны для падежа $selectedCase отсутствуют.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: диапазоны для падежа не найдены.")
        return
    }

    val currentState = userStates[chatId] ?: 0
    println("53. 🔎 Текущее состояние: $currentState")

    if (currentState >= rangesForCase.size) {
        println("54. ✅ Все этапы завершены для падежа: $selectedCase")
        addScoreForCase(chatId, selectedCase, filePath)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    val range = rangesForCase[currentState]
    println("55. 🔍 Генерация сообщения для диапазона: $range")
    val listName: String = "Существительные 2"
    val messageText = try {
        generateMessageFromRange(filePath, listName, range, wordUz, wordRus)
    } catch (e: Exception) {
        println("56. ❌ Ошибка при генерации сообщения: ${e.message}")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        return
    }

    println("57. ✅ Сообщение сгенерировано.")
    if (currentState == rangesForCase.size - 1) {
        println("571 $messageText")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2
        )
        println("58. ✅ Последний этап завершен. Добавляем балл и отправляем финальное меню.")
        val currentBlock = userBlocks[chatId] ?: 1
        println("59. ✅ Уточняем текущий блок пользователя.")
        addScoreForCase(chatId, selectedCase, filePath, currentBlock)
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
fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String, wordRus: String, filePath: String) {
    println("92. 🔍 Формируем меню финальных действий для пользователя $chatId")
    val showNextStep = checkUserState(chatId, filePath) // Проверяем состояние пользователя
    println("93. 🔎 Состояние пользователя. Добавить кнопку 'Следующий блок': $showNextStep")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Повторить", "repeat:$wordUz:$wordRus")),
        listOf(InlineKeyboardButton.CallbackData("Изменить слово", "change_word")),
        listOf(InlineKeyboardButton.CallbackData("Изменить падеж", "change_case")),
        listOf(InlineKeyboardButton.CallbackData("Изменить падеж и слово", "reset"))
    )

    if (showNextStep) {
        println("94. ✅ Добавляем кнопку 'Следующий блок'")
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Следующий блок", "next_block")))
    }

    println("95. 📨 Отправляем сообщение с финальным меню.")
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
}

fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String, wordRus: String): String {
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

    val firstCell = cells.firstOrNull()?.escapeMarkdownV2() ?: ""
    val blurredFirstCell = if (firstCell.isNotBlank()) "||$firstCell||" else ""
    println("101. 🔑 Первый элемент: \"$firstCell\", Заблюренный: \"$blurredFirstCell\"")

    val messageBody = cells.drop(1).joinToString("\n")
    println("102. 📄 Содержимое тела сообщения:\n$messageBody")

    workbook.close()
    println("103. 📕 Файл закрыт. Генерация завершена.")
    return listOf(blurredFirstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
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


fun adjustWordUz(content: String, wordUz: String): String {
    println("105. 🔍 Обработка слова \"$wordUz\" в контексте \"$content\"")

    if (content.contains("+")) {
        val replacement = "s" // Пример для логики `+`. Можешь оставить как раньше.
        println("106. 🔧 Найден '+'. Заменяем его на \"$replacement\".")
        return content.replace("+", replacement).replace("*", wordUz)
    }

    if (content.contains("*")) {
        println("107. 🔧 Найден '*'. Заменяем его на \"$wordUz\".")
        return content.replace("*", wordUz)
    }

    println("108. 🔎 Символов для замены не найдено.")
    return content
}

// extractCellsFromRange: Извлекает и обрабатывает ячейки из указанного диапазона листа Excel.
fun extractCellsFromRange(sheet: Sheet, range: String, wordUz: String): List<String> {
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

fun checkUserState(userId: Long, filePath: String, sheetName: String = "Состояние пользователя" ): Boolean {
    println("113. 🔍 Проверяем состояние пользователя. ID: $userId, Файл: $filePath, Лист: $sheetName")

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

            if (currentId == userId) {
                println("118. ✅ Пользователь найден в строке: ${rowIndex + 1}")
                userRow = row
                break
            }

            if (currentId == null || currentId == 0L) {
                println("119. ⚠️ Пустая ячейка. Записываем нового пользователя.")
                idCell.setCellValue(userId.toDouble())
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
            newRow.createCell(0).setCellValue(userId.toDouble())
            for (i in 1..6) {
                newRow.createCell(i).setCellValue(0.0)
            }
            safelySaveWorkbook(workbook, filePath)
            println("123. ✅ Пользователь добавлен в строку: ${newRowIndex + 1}")
            return false
        }
    }

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



fun addScoreForCase(userId: Long, case: String, filePath: String, sheetName: String = "Состояние пользователя") {
    println("163. 🔍 Начинаем добавление балла. Пользователь: $userId, Падеж: $case")

    val file = File(filePath)
    if (!file.exists()) {
        println("164. ❌ Файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден.")
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
            ?: throw IllegalArgumentException("165. ❌ Лист $sheetName не найден")
        println("166. ✅ Лист найден: $sheetName")

        val caseColumnIndex = when (case) {
            "Именительный" -> 1
            "Родительный" -> 2
            "Винительный" -> 3
            "Дательный" -> 4
            "Местный" -> 5
            "Исходный" -> 6
            else -> throw IllegalArgumentException("167. ❌ Неизвестный падеж: $case")
        }

        println("168. 🔎 Ищем пользователя $userId в таблице...")
        for (rowIndex in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue
            val idCell = row.getCell(0)
            val currentId = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            println("169. 🔎 Проверяем строку ${rowIndex + 1}: ID = $currentId.")
            if (currentId == userId) {
                println("170. ✅ Пользователь найден в строке ${rowIndex + 1}")
                val caseCell = row.getCell(caseColumnIndex) ?: row.createCell(caseColumnIndex)
                val currentScore = caseCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("171. 🔎 Текущее значение: $currentScore")
                caseCell.setCellValue(currentScore + 1)
                println("172. ✅ Новое значение: ${currentScore + 1}")
                safelySaveWorkbook(workbook, filePath)
                println("173. ✅ Балл добавлен и изменения сохранены.")
                return
            }
        }
        println("174. ⚠️ Пользователь $userId не найден. Этого не должно происходить на этом этапе.")
    }
}

fun checkUserState(userId: Long, filePath: String, sheetName: String = "Состояние пользователя", block: Int = 1): Boolean {
    println("175. 🔍 Проверка состояния пользователя. ID: $userId, файл: $filePath, лист: $sheetName, блок: $block.")

    val columnRanges = mapOf(
        1 to (1..6),
        2 to (7..12),
        3 to (13..18)
    )
    val columns = columnRanges[block] ?: run {
        println("176. ⚠️ Неизвестный блок: $block. Возвращаем false.")
        return false
    }

    println("177. ✅ Диапазон колонок для блока $block: $columns.")

    val file = File(filePath)
    if (!file.exists()) {
        println("178. ❌ Файл $filePath не найден. Возвращаем false.")
        return false
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            println("179. ❌ Лист $sheetName не найден. Возвращаем false.")
            return false
        }

        println("180. ✅ Лист $sheetName найден. Всего строк: ${sheet.lastRowNum + 1}.")

        for (row in sheet) {
            val idCell = row.getCell(0) ?: continue
            val userIdFromCell = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            println("181. 🔎 Строка ${row.rowNum + 1}: ID в ячейке = $userIdFromCell.")
            if (userIdFromCell == userId) {
                println("182. ✅ Пользователь найден в строке ${row.rowNum + 1}. Проверяем выполнение всех условий...")

                val allColumnsCompleted = columns.all { colIndex ->
                    val cell = row.getCell(colIndex)
                    val value = cell?.numericCellValue ?: 0.0
                    println("183. 📊 Колонка $colIndex: значение = $value. Выполнено? ${value > 0}")
                    value > 0
                }

                println("184. ✅ Результат проверки для пользователя $userId: $allColumnsCompleted.")
                return allColumnsCompleted
            }
        }
    }

    println("185. ⚠️ Пользователь $userId не найден в таблице. Возвращаем false.")
    return false
}


fun addScoreForCase(userId: Long, case: String, filePath: String, block: Int) {
    println("186. 🔍 Начинаем добавление балла для пользователя $userId. Падеж: $case, Файл: $filePath, Блок: $block.")

    val columnRanges = mapOf(
        1 to mapOf("Именительный" to 1, "Родительный" to 2, "Винительный" to 3, "Дательный" to 4, "Местный" to 5, "Исходный" to 6),
        2 to mapOf("Именительный" to 7, "Родительный" to 8, "Винительный" to 9, "Дательный" to 10, "Местный" to 11, "Исходный" to 12),
        3 to mapOf("Именительный" to 13, "Родительный" to 14, "Винительный" to 15, "Дательный" to 16, "Местный" to 17, "Исходный" to 18)
    )

    val column = columnRanges[block]?.get(case)
    if (column == null) {
        println("187. ❌ Ошибка: колонка для блока $block и падежа $case не найдена.")
        return
    }
    println("188. ✅ Определена колонка для записи: $column.")

    val file = File(filePath)
    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
        if (sheet == null) {
            println("189. ❌ Ошибка: Лист 'Состояние пользователя' не найден.")
            return
        }
        println("190. ✅ Лист 'Состояние пользователя' найден. Всего строк в листе: ${sheet.lastRowNum + 1}.")

        var userFound = false

        for (row in sheet) {
            val idCell = row.getCell(0)
            if (idCell == null) {
                println("191. ⚠️ Ячейка ID отсутствует в строке ${row.rowNum + 1}, пропускаем.")
                continue
            }

            val idFromCell = try {
                idCell.numericCellValue.toLong()
            } catch (e: Exception) {
                println("192. ❌ Ошибка: не удалось прочитать ID из ячейки в строке ${row.rowNum + 1}. Значение: ${idCell}.")
                continue
            }

            println("193. 🔎 Проверяем строку ${row.rowNum + 1}. ID в ячейке: $idFromCell.")

            if (idFromCell == userId) {
                userFound = true
                println("194. ✅ Пользователь найден в строке ${row.rowNum + 1}. Начинаем обновление.")
                val targetCell = row.getCell(column) ?: row.createCell(column)
                val currentValue = targetCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("195. 🔎 Текущее значение в колонке $column: $currentValue.")

                targetCell.setCellValue(currentValue + 1)
                println("196. ✅ Новое значение в колонке $column: ${currentValue + 1}. Сохраняем файл.")

                safelySaveWorkbook(workbook, filePath)
                println("197. ✅ Изменения сохранены в файле $filePath.")
                return
            }
        }
        if (!userFound) {
            println("198. ⚠️ Пользователь с ID $userId не найден. Новая запись не создана.")
        }
    }
}

fun calculateTotalScore(userId: Long, filePath: String, sheetName: String = "Состояние пользователя"): Int {
    println("199. 🔍 Начало вычисления общего счёта для пользователя $userId в файле $filePath, лист $sheetName.")
    val columnIndices = (1..18) // Все столбцы, относящиеся к блокам (B-S)
    val file = File(filePath)

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            println("200. ❌ Лист $sheetName не найден. Возвращаем 0.")
            return 0
        }

        println("201. ✅ Лист $sheetName найден. Начинаем обход строк.")
        for (row in sheet) {
            val idCell = row.getCell(0)
            val userIdFromCell = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            println("202. 🔎 Строка ${row.rowNum + 1}: ID пользователя в ячейке = $userIdFromCell.")
            if (userIdFromCell == userId) {
                val totalScore = columnIndices.sumOf { colIndex ->
                    val cell = row.getCell(colIndex)
                    val value = (cell?.numericCellValue ?: 0.0).toInt()
                    println("203. 📊 Колонка $colIndex: значение = $value.")
                    value
                }
                println("204. ✅ Общий счёт для пользователя $userId: $totalScore.")
                return totalScore
            }
        }
    }

    println("205. ⚠️ Пользователь $userId не найден в таблице. Возвращаем 0.")
    return 0
}

// processCellContent: Обрабатывает содержимое ячейки с учётом цвета текста и символа `+`.
fun processCellContent(cell: Cell?, wordUz: String): String {
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

    // Если форматированных участков нет, обрабатываем текст целиком
    if (runs == 0) {
        println("⚠️ У ячейки нет форматированных участков. Используем весь текст.")
        val processedContent = adjustWordUz(text, wordUz).escapeMarkdownV2()
        println("✅ Результат для текста без форматирования: \"$processedContent\"")
        return processedContent
    }

    // Если форматированные участки есть, обрабатываем их по частям
    val processedText = buildString {
        for (i in 0 until runs) {
            val start = richText.getIndexOfFormattingRun(i)
            val end = if (i + 1 < runs) richText.getIndexOfFormattingRun(i + 1) else text.length
            val substring = text.substring(start, end)

            val font = richText.getFontOfFormattingRun(i) as XSSFFont?
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
    }

    println("213. ✅ Результат обработки: \"$processedText\".")
    return processedText
}



// Функция для извлечения цвета шрифта
fun getFontColor(font: XSSFFont): String {
    val color = font.xssfColor
    val colorHex = color?.rgb?.joinToString(prefix = "#", separator = "") { "%02X".format(it) } ?: "Цвет не определён"
    println("        🔍 Цвет шрифта: $colorHex")
    return colorHex
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






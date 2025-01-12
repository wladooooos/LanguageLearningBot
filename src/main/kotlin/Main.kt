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
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import java.io.FileOutputStream

val caseRanges = mapOf(
    "Именительный" to listOf("A1-A4", "B1-B7", "C1-C7", "D1-D7", "E1-E7"),
    "Родительный" to listOf("A8-A9", "B8-B14", "C8-C14", "D8-D14", "E8-E14"),
    "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21"),
    "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28"),
    "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35"),
    "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42")
)

val userStates = mutableMapOf<Long, Int>()
val userCases = mutableMapOf<Long, String>() // Хранение выбранного падежа для каждого пользователя
val userWords = mutableMapOf<Long, Pair<String, String>>() // Хранение выбранного слова для каждого пользователя
val tableFile = "Алгоритм 3.3.xlsx"

fun main() {

    val bot = bot {
        token = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE"

        dispatch {
            command("start") {
                val userId = message.chat.id
                println("🔍 Команда /start от пользователя: $userId")
                userStates[userId] = 0
                userWords.remove(userId) // Удаляем старые данные о слове
                sendWelcomeMessage(userId, bot)
                sendCaseSelection(userId, bot, tableFile) // Добавляем tableFile как filePath
            }

            callbackQuery {
                val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery
                val data = callbackQuery.data ?: return@callbackQuery
                println("🔍 Получен callback от пользователя: $chatId, данные: $data")

                println("🔍 Обработка callbackQuery: $data для пользователя $chatId")

                when {
                    data.startsWith("case:") -> {
                        val selectedCase = data.removePrefix("case:")
                        println("✅ Выбранный падеж: $selectedCase")
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
                        println("🔍 Выбор слова: $data")
                        val (wordUz, wordRus) = extractWordsFromCallback(data)
                        println("✅ Выбранное слово: $wordUz ($wordRus)")
                        userWords[chatId] = Pair(wordUz, wordRus)
                        userStates[chatId] = 0
                        sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                    }
                    data.startsWith("next:") -> {
                        println("🔍 Обработка кнопки 'Далее': $data")
                        val params = data.removePrefix("next:").split(":")
                        if (params.size == 2) {
                            val wordUz = params[0]
                            val wordRus = params[1]
                            val currentState = userStates[chatId] ?: 0
                            println("🔎 Текущее состояние: $currentState. Обновляем.")
                            userStates[chatId] = currentState + 1
                            sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                        }
                    }
                    data.startsWith("repeat:") -> {
                        println("🔍 Обработка кнопки 'Повторить': $data")
                        val params = data.removePrefix("repeat:").split(":")
                        if (params.size == 2) {
                            val wordUz = params[0]
                            val wordRus = params[1]
                            userStates[chatId] = 0
                            sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                        }
                    }
                    data == "change_word" -> {
                        println("🔍 Выбор нового слова для пользователя: $chatId")
                        userStates[chatId] = 0
                        sendWordMessage(chatId, bot, tableFile)
                    }
                    data == "change_case" -> {
                        println("🔍 Выбор нового падежа для пользователя: $chatId")
                        sendCaseSelection(chatId, bot, tableFile)
                    }
                    data == "reset" -> {
                        println("🔍 Полный сброс данных для пользователя: $chatId")
                        userWords.remove(chatId)
                        userCases.remove(chatId)
                        userStates.remove(chatId)
                        sendCaseSelection(chatId, bot, tableFile)
                    }
                    data == "test" -> {
                        println("🔍 Запрос на тестовую функцию от пользователя: $chatId")
                        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Тест в разработке.")
                    }
                    else -> {
                        println("⚠️ Неизвестный callbackQuery: $data")
                    }
                }
            }
        }
    }

    bot.startPolling()
}
// sendCaseSelection: Отправляет пользователю клавиатуру для выбора падежа.
fun sendCaseSelection(chatId: Long, bot: Bot, filePath: String) {
    println("🔍 Формируем меню выбора падежей для пользователя: $chatId")
    val showNextStep = checkUserState(chatId, filePath)
    println("🔎 Проверка состояния завершена. Добавить кнопку 'Следующий шаг': $showNextStep")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Именительный падеж", "case:Именительный")),
        listOf(InlineKeyboardButton.CallbackData("Родительный падеж", "case:Родительный")),
        listOf(InlineKeyboardButton.CallbackData("Винительный падеж", "case:Винительный")),
        listOf(InlineKeyboardButton.CallbackData("Дательный падеж", "case:Дательный")),
        listOf(InlineKeyboardButton.CallbackData("Местный падеж", "case:Местный")),
        listOf(InlineKeyboardButton.CallbackData("Исходный падеж", "case:Исходный"))
    )

    if (showNextStep) {
        println("✅ Добавляем кнопку 'Следующий шаг'")
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Следующий шаг", "next_step")))
    }

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите падеж для изучения:",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
}



// sendWelcomeMessage: Отправляет приветственное сообщение с кнопкой /start.
fun sendWelcomeMessage(chatId: Long, bot: Bot) {
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
    return listOf(
        listOf(KeyboardButton("/start"))
    )
}

// sendWordMessage: Отправляет клавиатуру с выбором слов.
fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
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

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите слово из списка:",
        replyMarkup = inlineKeyboard
    )
}

// createWordSelectionKeyboardFromExcel: Создает клавиатуру из данных в Excel-файле.
fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
    println("🔍 Начало создания клавиатуры из файла: $filePath, лист: $sheetName")

    // Проверка существования файла
    println("🔎 Проверяем существование файла $filePath")
    val file = File(filePath)
    if (!file.exists()) {
        println("❌ Файл $filePath не найден")
        throw IllegalArgumentException("Файл $filePath не найден")
    }
    println("✅ Файл найден: $filePath")

    // Открытие файла и получение листа
    println("📂 Открываем файл Excel")
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName)
        ?: throw IllegalArgumentException("❌ Лист $sheetName не найден")
    println("✅ Лист найден: $sheetName")

    println("📜 Всего строк в листе: ${sheet.lastRowNum + 1}")

    // Генерация случайных строк
    println("🎲 Генерация случайных строк для кнопок")
    val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
    println("🎲 Выбраны строки: $randomRows")

    // Создание кнопок
    println("🔨 Начинаем обработку строк для создания кнопок")
    val buttons = randomRows.mapNotNull { rowIndex ->
        val row = sheet.getRow(rowIndex)
        if (row == null) {
            println("⚠️ Строка $rowIndex отсутствует, пропускаем")
            return@mapNotNull null
        }

        val wordUz = row.getCell(0)?.toString()?.trim()
        val wordRus = row.getCell(1)?.toString()?.trim()

        if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
            println("⚠️ Некорректные данные в строке $rowIndex: wordUz = $wordUz, wordRus = $wordRus. Пропускаем")
            return@mapNotNull null
        }

        println("✅ Обработана строка $rowIndex: wordUz = $wordUz, wordRus = $wordRus")
        InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
    }.chunked(2) // Группируем кнопки по 2 в строке

    println("✅ Кнопки успешно созданы. Количество строк кнопок: ${buttons.size}")

    // Закрытие файла
    println("📕 Закрываем файл Excel")
    workbook.close()
    println("✅ Файл Excel успешно закрыт")

    // Генерация клавиатуры завершена
    println("🔑 Генерация клавиатуры завершена")
    return InlineKeyboardMarkup.create(buttons)
}

// extractWordsFromCallback: Извлекает узбекское и русское слово из callback data.
fun extractWordsFromCallback(data: String): Pair<String, String> {
    val content = data.substringAfter("word:(").substringBefore(")")
    val wordUz = content.substringBefore(" - ").trim()
    val wordRus = content.substringAfter(" - ").trim()
    return wordUz to wordRus
}

// sendStateMessage: Отправляет сообщение по текущему состоянию и падежу.
fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String, wordRus: String) {
    println("🔍 Формируем сообщение для пользователя: $chatId, слово: $wordUz, перевод: $wordRus")

    val selectedCase = userCases[chatId]
    if (selectedCase == null) {
        println("❌ Ошибка: выбранный падеж отсутствует.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: выберите падеж.")
        return
    }

    println("✅ Выбранный падеж: $selectedCase")

    val rangesForCase = caseRanges[selectedCase]
    if (rangesForCase == null) {
        println("❌ Ошибка: диапазоны для падежа $selectedCase отсутствуют.")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка: диапазоны для падежа не найдены.")
        return
    }

    val currentState = userStates[chatId] ?: 0
    println("🔎 Текущее состояние: $currentState")

    if (currentState >= rangesForCase.size) {
        println("✅ Все этапы завершены для падежа: $selectedCase")
        addScoreForCase(chatId, selectedCase, filePath)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    val range = rangesForCase[currentState]
    println("🔍 Генерация сообщения для диапазона: $range")

    val messageText = try {
        generateMessageFromRange(filePath, "Примеры гамм для существительны", range, wordUz, wordRus)
    } catch (e: Exception) {
        println("❌ Ошибка при генерации сообщения: ${e.message}")
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        return
    }

    println("✅ Сообщение сгенерировано.")
    if (currentState == rangesForCase.size - 1) {
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2
        )
        println("✅ Последний этап завершен. Добавляем балл и отправляем финальное меню.")
        addScoreForCase(chatId, selectedCase, filePath)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
    } else {
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
    println("🔍 Формируем меню финальных действий для пользователя $chatId")
    val showNextStep = checkUserState(chatId, filePath) // Проверяем состояние пользователя
    println("🔎 Состояние пользователя. Добавить кнопку 'Следующий шаг': $showNextStep")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Повторить", "repeat:$wordUz:$wordRus")),
        listOf(InlineKeyboardButton.CallbackData("Изменить слово", "change_word")),
        listOf(InlineKeyboardButton.CallbackData("Изменить падеж", "change_case")),
        listOf(InlineKeyboardButton.CallbackData("Изменить падеж и слово", "reset"))
    )

    // Добавляем кнопку "Следующий шаг", если все условия выполнены
    if (showNextStep) {
        println("✅ Добавляем кнопку 'Следующий шаг'")
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Следующий шаг", "next_step")))
    }

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
}

// generateMessageFromRange: Генерирует текст сообщения из указанного диапазона ячеек Excel.
fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String, wordRus: String): String {
    val file = File(filePath)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName)
        ?: throw IllegalArgumentException("Лист $sheetName не найден")
    val cells = extractCellsFromRange(sheet, range)

    val firstCell = cells.firstOrNull()?.toString()?.replace("\n\n", "\n")?.trim()?.escapeMarkdownV2() ?: ""
    val blurredFirstCell = if (firstCell.isNotBlank()) "||$firstCell||" else ""

    val specificRowsForE = setOf(8, 9, 12, 13, 18, 19, 20, 25, 26, 27, 31, 32, 36, 38, 40)
    val messageBody = cells.drop(1).joinToString("\n") { cell ->
        val content = cell.toString().replace("*", wordUz)
        if (cell.columnIndex == 2 || (cell.columnIndex == 4 && specificRowsForE.contains(cell.rowIndex))) {
            adjustWordUz(content, wordUz)
        } else {
            content
        }.replace("`", "\\`")
    }

    workbook.close()
    return listOf(blurredFirstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
}

// escapeMarkdownV2: Экранирует спецсимволы для Markdown.
fun String.escapeMarkdownV2(): String {
    return this.replace("\\", "\\\\")
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
}

// adjustWordUz: Модифицирует узбекское слово в зависимости от контекста.
fun adjustWordUz(content: String, wordUz: String): String {
    val vowels = "aeiouаоиеёу"
    val lastChar = wordUz.lastOrNull()?.lowercaseChar()
    val nextChar = content.substringAfter("+", "").firstOrNull()?.lowercaseChar()

    if (!content.contains("+")) {
        return content
    }

    val replacement = when {
        lastChar != null && nextChar != null &&
                lastChar in vowels && nextChar in vowels -> "s"
        lastChar != null && nextChar != null &&
                lastChar !in vowels && nextChar !in vowels -> "i"
        else -> ""
    }

    return content.replace("+", replacement).replace("*", wordUz)
}

// extractCellsFromRange: Извлекает ячейки из указанного диапазона листа Excel.
fun extractCellsFromRange(sheet: Sheet, range: String): List<Cell> {
    val (start, end) = range.split("-").map { it.replace(Regex("[A-Z]"), "").toInt() - 1 }
    val column = range[0] - 'A'

    return (start..end).mapNotNull { rowIndex ->
        val row = sheet.getRow(rowIndex)
        row?.getCell(column)
    }
}

fun checkUserState(userId: Long, filePath: String, sheetName: String = "Состояние пользователя"): Boolean {
    println("🔍 Проверяем состояние пользователя. ID: $userId, Файл: $filePath, Лист: $sheetName")

    val file = File(filePath)
    var userRow: Row? = null
    var allCompleted = false

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден")
        println("✅ Лист найден: $sheetName")

        // Ищем ID пользователя или первую пустую строку
        println("🔍 Начинаем поиск пользователя по ID или пустой строки")
        for (rowIndex in 1..sheet.lastRowNum.coerceAtLeast(1)) { // Пропускаем первую строку (заголовки)
            val row = sheet.getRow(rowIndex) ?: sheet.createRow(rowIndex) // Создаём строку, если её нет
            val idCell = row.getCell(0) ?: row.createCell(0) // Создаём ячейку ID, если её нет

            // Чтение значения ID
            val currentId = when (idCell.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong() // Приводим Double к Long
                CellType.STRING -> idCell.stringCellValue.toLongOrNull() // Пробуем привести строку к Long
                else -> null
            }
            println("🔎 Проверяем строку: ${rowIndex + 1}, ID: $currentId")

            // Если ID совпадает, проверяем остальные ячейки
            if (currentId == userId) {
                println("✅ Пользователь найден в строке: ${rowIndex + 1}")
                userRow = row
                break
            }

            // Если ID отсутствует, записываем нового пользователя
            if (currentId == null || currentId == 0L) {
                println("⚠️ Пустая ячейка. Записываем нового пользователя.")
                idCell.setCellValue(userId.toDouble()) // Явно записываем как Double
                for (i in 1..6) {
                    row.createCell(i).setCellValue(0.0) // Устанавливаем значения по умолчанию
                }
                safelySaveWorkbook(workbook, filePath) // Сохраняем изменения
                println("✅ Пользователь добавлен в строку: ${rowIndex + 1}")
                return false // Новому пользователю доступ к "Следующему шагу" не предоставляется
            }
        }

        // Если ID найден, проверяем остальные ячейки
        if (userRow != null) {
            println("🔍 Проверяем значения в колонках Им., Род., Вин., Дат., Мест., Исх.")
            allCompleted = (1..6).all { index ->
                val cell = userRow!!.getCell(index)
                val value = cell?.toString()?.toDoubleOrNull() ?: 0.0
                println("🔎 Колонка ${index + 1}: значение = $value")
                value > 0
            }
        } else {
            // Если мы прошли весь список и не нашли ID, добавляем его в новую строку
            val newRowIndex = sheet.physicalNumberOfRows
            val newRow = sheet.createRow(newRowIndex)
            newRow.createCell(0).setCellValue(userId.toDouble()) // ID в первом столбце
            for (i in 1..6) {
                newRow.createCell(i).setCellValue(0.0) // Значения по умолчанию
            }
            safelySaveWorkbook(workbook, filePath) // Сохраняем изменения
            println("✅ Пользователь добавлен в строку: ${newRowIndex + 1}")
            return false // Новому пользователю доступ к "Следующему шагу" не предоставляется
        }
    }

    println("✅ Проверка завершена. Все шаги завершены: $allCompleted")
    return allCompleted
}



fun safelySaveWorkbook(workbook: org.apache.poi.ss.usermodel.Workbook, filePath: String) {
    val tempFile = File("${filePath}.tmp")
    try {
        FileOutputStream(tempFile).use { outputStream ->
            workbook.write(outputStream)
        }
        val originalFile = File(filePath)
        if (originalFile.exists()) {
            if (!originalFile.delete()) {
                throw IllegalStateException("Не удалось удалить старый файл: $filePath")
            }
        }
        if (!tempFile.renameTo(originalFile)) {
            throw IllegalStateException("Не удалось переименовать временный файл: ${tempFile.path}")
        }
        println("✅ Файл успешно сохранён: $filePath")
    } catch (e: Exception) {
        println("❌ Ошибка при сохранении файла: ${e.message}")
        tempFile.delete() // Удаляем временный файл в случае ошибки
        throw e
    }
}

fun addScoreForCase(userId: Long, case: String, filePath: String, sheetName: String = "Состояние пользователя") {
    println("🔍 Начинаем добавление балла. Пользователь: $userId, Падеж: $case")

    val file = File(filePath)
    if (!file.exists()) {
        println("❌ Файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден.")
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("❌ Лист $sheetName не найден")
        println("✅ Лист найден: $sheetName")

        // Определяем индекс столбца для падежа
        val caseColumnIndex = when (case) {
            "Именительный" -> 1
            "Родительный" -> 2
            "Винительный" -> 3
            "Дательный" -> 4
            "Местный" -> 5
            "Исходный" -> 6
            else -> throw IllegalArgumentException("❌ Неизвестный падеж: $case")
        }

        println("🔎 Ищем пользователя $userId в таблице...")
        for (rowIndex in 1..sheet.lastRowNum) { // Пропускаем заголовок
            val row = sheet.getRow(rowIndex) ?: continue
            val idCell = row.getCell(0)
            val currentId = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong() // Приводим Double к Long
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong() // Сначала в Double, затем в Long
                else -> null
            }

            if (currentId == userId) {
                println("✅ Пользователь найден в строке ${rowIndex + 1}")
                val caseCell = row.getCell(caseColumnIndex) ?: row.createCell(caseColumnIndex)
                val currentScore = caseCell.numericCellValue.takeIf { it > 0 } ?: 0.0
                println("🔎 Текущее значение: $currentScore")
                caseCell.setCellValue(currentScore + 1)
                println("✅ Новое значение: ${currentScore + 1}")
                safelySaveWorkbook(workbook, filePath)
                println("✅ Балл добавлен и изменения сохранены.")
                return
            }
        }
        println("⚠️ Пользователь $userId не найден. Этого не должно происходить на этом этапе.")
    }
}

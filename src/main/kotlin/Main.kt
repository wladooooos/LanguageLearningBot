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
import jdk.internal.org.jline.utils.InfoCmp.Capability.buttons

val stateRanges = listOf(
    "A1-A4", "B1-B7", "C1-C7", "D1-D7", "E1-E7",
    "A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14",
    "A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21",
    "A22-A23", "B22-B23", "C22-C28", "D22-D28", "E22-E28",
    "A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35",
    "A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42"
)

val userStates = mutableMapOf<Long, Int>()

fun main() {
    println("1. Начало работы программы")
    val tableFile = "Алгоритм 2.13.xlsx"
    println("2. Указан файл таблицы: $tableFile")

    val bot = bot {
        token = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE"
        println("3. Токен бота установлен")

        dispatch {
            println("4. Настройка обработчиков команд и коллбеков")
            command("start") {
                println("5. Получена команда /start от пользователя ${message.chat.id}")

                // Установка начального состояния пользователя
                val userId = message.chat.id
                println("5.1 Устанавливаем начальное состояние для пользователя $userId")
                userStates[userId] = 0
                println("6. Установлено начальное состояние для пользователя $userId: ${userStates[userId]}")

                // Отправка приветственного сообщения
                println("6.1 Отправляем приветственное сообщение для пользователя $userId")
                sendWelcomeMessage(userId, bot)
                println("6.2 Приветственное сообщение успешно отправлено для пользователя $userId")

                // Отправка сообщения с выбором слов
                println("6.3 Отправляем сообщение с выбором слов для пользователя $userId")
                sendWordMessage(userId, bot, tableFile)
                println("6.4 Сообщение с выбором слов успешно отправлено для пользователя $userId")
            }

            callbackQuery {
                val chatId = callbackQuery.message?.chat?.id
                if (chatId == null) {
                    println("7. Ошибка: chatId отсутствует в callbackQuery")
                    return@callbackQuery
                }
                val data = callbackQuery.data ?: return@callbackQuery
                println("8. Получен callbackQuery от пользователя $chatId с данными: $data")

                when {
                    data.startsWith("word:") -> {
                        val (wordUz, wordRus) = extractWordsFromCallback(data)
                        println("9. Извлечены слова из callback: wordUz = $wordUz, wordRus = $wordRus")
                        userStates[chatId] = 0
                        println("10. Сброшено состояние для пользователя $chatId: ${userStates[chatId]}")
                        sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                    }
                    data.startsWith("next:") -> {
                        val params = data.removePrefix("next:").split(":")
                        if (params.size == 2) {
                            val wordUz = params[0]
                            val wordRus = params[1]
                            println("Извлечён коллбэк: wordUz = $wordUz, wordRus = $wordRus")

                            // Обновление состояния и отправка следующего сообщения
                            val currentState = userStates[chatId] ?: 0
                            userStates[chatId] = currentState + 1
                            println("Обновлено состояние для пользователя $chatId: ${userStates[chatId]}")
                            sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
                        } else {
                            println("Ошибка: недостаточно параметров в коллбэке")
                        }
                    }
                    data == "test" -> {
                        println("16. Пользователь $chatId выбрал тест")
                        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Тест в разработке.")
                    }
                    data == "new_word" -> {
                        println("5.1 Устанавливаем начальное состояние для пользователя $chatId")
                        userStates[chatId] = 0
                        println("6. Установлено начальное состояние для пользователя $chatId: ${userStates[chatId]}")
                        // Отправка сообщения с выбором слов
                        println("6.3 Отправляем сообщение с выбором слов для пользователя $chatId")
                        sendWordMessage(chatId, bot, tableFile)
                        println("6.4 Сообщение с выбором слов успешно отправлено для пользователя $chatId")
                    }
                }
            }
        }
    }

    bot.startPolling()
    println("17. Бот начал опрос")
}

//  Отправляет приветственное сообщение с клавиатурой пользователю.
fun sendWelcomeMessage(chatId: Long, bot: Bot) {
    println("40. Отправка приветственного сообщения и пользовательской клавиатуры пользователю $chatId")
    val keyboardMarkup = KeyboardReplyMarkup( // Создаём клавиатуру с кнопками пользователя
        keyboard = generateUsersButton(),     // Вызываем функцию для генерации кнопок
        resizeKeyboard = true                 // Указываем, что клавиатура должна автоматически подстраиваться под экран
    )
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = """Здравствуйте!
Я бот, помогающий изучать узбекский язык!
                    """.trimIndent(),
        replyMarkup = keyboardMarkup
    )
    println("41. Приветственное сообщение отправлено пользователю $chatId")
}

// Генерирует кнопки для пользовательской клавиатуры.
fun generateUsersButton(): List<List<KeyboardButton>> {
    return listOf(
        listOf(KeyboardButton("/start"))
    )
}

// Отправляет пользователю сообщение с выбором слов, загружая данные из Excel-файла.
fun sendWordMessage(chatId: Long, bot: Bot, filePath: String) {
    println("18. Начинаем отправку меню со словами для пользователя $chatId")

    // Проверяем наличие файла
    println("18.1 Проверяем существование файла: $filePath")
    if (!File(filePath).exists()) {
        println("18.2 Ошибка: Файл $filePath не найден")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка: файл с данными не найден."
        )
        return
    }

    // Генерация клавиатуры
    println("18.3 Генерируем клавиатуру из файла $filePath")
    val inlineKeyboard = try {
        createWordSelectionKeyboardFromExcel(filePath, "Существительные")
    } catch (e: Exception) {
        println("18.4 Ошибка при создании клавиатуры: ${e.message}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка при обработке данных: ${e.message}"
        )
        return
    }
    println("18.5 Клавиатура успешно сгенерирована")

    // Отправка сообщения с клавиатурой
    println("18.6 Отправляем сообщение с выбором слов пользователю $chatId")
    try {
        println("18.7 Клавиатура для отправки: ${inlineKeyboard.inlineKeyboard.flatten()}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Выберите слово из списка:",
            replyMarkup = inlineKeyboard
        )
        println("19. Сообщение с выбором слов успешно отправлено пользователю $chatId")
    } catch (e: Exception) {
        println("18.7 Ошибка при отправке сообщения: ${e.message}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка при отправке сообщения: ${e.message}"
        )
    }
}

// Генерирует клавиатуру из случайных строк указанного листа Excel.
fun createWordSelectionKeyboardFromExcel(filePath: String, sheetName: String): InlineKeyboardMarkup {
    println("20. Начинаем генерацию клавиатуры из файла: $filePath, лист: $sheetName")

    // Проверка существования файла
    println("20.1 Проверяем существование файла $filePath")
    val file = File(filePath)
    if (!file.exists()) {
        throw IllegalArgumentException("Файл $filePath не найден")
    }

    // Открытие файла и проверка листа
    println("20.2 Открываем файл Excel")
    val workbook = try {
        WorkbookFactory.create(file)
    } catch (e: Exception) {
        println("20.3 Ошибка при открытии файла: ${e.message}")
        throw IllegalArgumentException("Ошибка при открытии файла Excel: ${e.message}")
    }
    println("20.4 Файл Excel успешно открыт")

    val sheet = workbook.getSheet(sheetName)
    if (sheet == null) {
        workbook.close()
        throw IllegalArgumentException("Лист $sheetName не найден")
    }
    println("21. Лист $sheetName найден, всего строк: ${sheet.lastRowNum}")

    // Выбор случайных строк
    val randomRows = (1..sheet.lastRowNum).shuffled().take(10)
    println("22. Выбраны случайные строки: $randomRows")

    // Создание кнопок
    val buttons = randomRows.mapNotNull { rowIndex ->
        println("22.1 Обрабатываем строку $rowIndex")
        val row = sheet.getRow(rowIndex)
        if (row == null) {
            println("22.2 Строка $rowIndex отсутствует, пропускаем")
            return@mapNotNull null
        }

        val wordUz = row.getCell(0)?.toString()?.trim()
        val wordRus = row.getCell(1)?.toString()?.trim()
        println("23. Обработаны данные строки $rowIndex: wordUz = $wordUz, wordRus = $wordRus")

        if (!wordUz.isNullOrBlank() && !wordRus.isNullOrBlank()) {
            InlineKeyboardButton.CallbackData("$wordUz - $wordRus", "word:($wordUz - $wordRus)")
        } else {
            println("23.1 Данные строки $rowIndex некорректны, пропускаем")
            null
        }
    }.chunked(2)

    // Закрытие файла
    println("24. Закрываем файл Excel")
    workbook.close()
    println("24.1 Файл Excel успешно закрыт")

    // Генерация клавиатуры завершена
    println("24.2 Генерация клавиатуры завершена, количество кнопок: ${buttons.size}")
    return InlineKeyboardMarkup.create(buttons)
}

// Извлекает слова из данных callback-запроса.
fun extractWordsFromCallback(data: String): Pair<String, String> {
    println("25. Извлечение слов из callback: $data")
    val content = data.substringAfter("word:(").substringBefore(")")
    val wordUz = content.substringBefore(" - ").trim()
    val wordRus = content.substringAfter(" - ").trim()
    println("26. Извлеченные слова: wordUz = $wordUz, wordRus = $wordRus")
    return wordUz to wordRus
}

// Отправляет сообщение о текущем состоянии на основе диапазона данных Excel.
fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String, wordRus: String) {
    println("27. Отправка сообщения для пользователя $chatId, состояние: ${userStates[chatId]}")
    val currentState = userStates[chatId] ?: 0
    val range = if (currentState < stateRanges.size) stateRanges[currentState] else null

    if (currentState == 29) {
        println("Достигнуто предпоследнее состояние. Отправляем два сообщения.")

        // Отправка первого сообщения без кнопки "Далее"
        val messageText = try {
            generateMessageFromRange(filePath, "Примеры гамм для существительны", range ?: "", wordUz, wordRus)
        } catch (e: Exception) {
            println("28.1 Ошибка при генерации сообщения: ${e.message}")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка при формировании сообщения: ${e.message}"
            )
            return
        }

        println("28.2 Сформированное сообщение (предпоследнее):\n$messageText")
        try {
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = messageText,
                parseMode = ParseMode.MARKDOWN_V2 // Без кнопки "Далее"
            )
            println("29.1 Предпоследнее сообщение успешно отправлено пользователю $chatId")
        } catch (e: Exception) {
            println("28.4 Ошибка при отправке предпоследнего сообщения: ${e.message}")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Ошибка при отправке предпоследнего сообщения: ${e.message}"
            )
            return
        }

        // Отправка второго сообщения с финальными кнопками
        sendFinalButtons(chatId, bot, wordUz, wordRus)
        return
    }

    if (range == null) {
        println("Достигнуто последнее состояние. Отправка финального сообщения.")
        sendFinalButtons(chatId, bot, wordUz, wordRus)
        return
    }

    println("28. Диапазон текущего состояния: $range")

    // Генерация сообщения
    val messageText = try {
        generateMessageFromRange(filePath, "Примеры гамм для существительны", range, wordUz, wordRus)
    } catch (e: Exception) {
        println("28.1 Ошибка при генерации сообщения: ${e.message}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка при формировании сообщения: ${e.message}"
        )
        return
    }

    println("28.2 Сформированное сообщение:\n$messageText")
    println("28.3 Параметры для отправки: chatId = $chatId, wordUz = $wordUz, wordRus = $wordRus")

    // Отправка сообщения
    try {
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2,
            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus")
            )
        )
        println("29. Сообщение успешно отправлено пользователю $chatId")
    } catch (e: Exception) {
        println("28.4 Ошибка при отправке сообщения: ${e.message}")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка при отправке сообщения: ${e.message}"
        )
    }
}

// Отправка последнего сообщения с тремя кнопками
fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String, wordRus: String) {
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим словом. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(
            listOf(
                listOf(InlineKeyboardButton.CallbackData("Повторить со словом \"$wordUz\"", "word:($wordUz - $wordRus)")),
                listOf(InlineKeyboardButton.CallbackData("Выбрать новое слово", "new_word")),
                listOf(InlineKeyboardButton.CallbackData("Пройти тест", "test"))
            )
        )
    )
}


// Генерирует сообщение из диапазона ячеек Excel и обрабатывает данные для замены символов.
fun generateMessageFromRange(filePath: String, sheetName: String, range: String, wordUz: String, wordRus: String): String {
    println("30. Генерация сообщения из диапазона: $range")
    val file = File(filePath)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден")
    val cells = extractCellsFromRange(sheet, range)
    println("31. Найдено ${cells.size} ячеек в диапазоне $range")

    // Обработка первой ячейки (справочная информация)
    val firstCell = cells.firstOrNull()?.toString()?.replace("\n\n", "\n")?.trim()?.escapeMarkdownV2() ?: ""
    val blurredFirstCell = if (firstCell.isNotBlank()) "||$firstCell||" else ""
    println("31.1 Скрытый текст первой ячейки: $blurredFirstCell")

    // Обработка остальных ячеек
    val specificRowsForE = setOf(8, 9, 12, 13, 18, 19, 20, 25, 26, 27, 31, 32, 36, 38, 40) // Индексы строк для ячеек E9, E10 и т.д.
    val messageBody = cells.drop(1).joinToString("\n") { cell ->
        val content = cell.toString().replace("*", wordUz)
        println("34.0 Исходное содержимое ячейки: $content")

        val adjustedContent = if (
            cell.columnIndex == 2 || // Для столбца C
            (cell.columnIndex == 4 && specificRowsForE.contains(cell.rowIndex)) // Для конкретных ячеек столбца E
        ) {
            adjustWordUz(content, wordUz)
        } else {
            content
        }

        println("34.2 Результат после обработки: $adjustedContent")
        adjustedContent.replace("`", "\\`") // Экранирование обратного апострофа
    }
    println("32. Текст сообщения без первой ячейки:\n$messageBody")

    workbook.close()
    println("33. Файл закрыт")

    // Формируем итоговое сообщение
    val finalMessage = listOf(blurredFirstCell, messageBody).filter { it.isNotBlank() }.joinToString("\n\n")
    println("33.1 Итоговое сообщение:\n$finalMessage")
    return finalMessage
}

// Экранирует символы для корректной работы формата MarkdownV2.
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

// Изменяет слово на основе контекста, заменяя "+" на буквы (s/i) для согласных и гласных.
fun adjustWordUz(content: String, wordUz: String): String {
    val vowels = "aeiouаоиеёу"
    val lastChar = wordUz.lastOrNull()?.lowercaseChar()
    val nextChar = content.substringAfter("+", "").firstOrNull()?.lowercaseChar()

    println("34. Обработка окончания слова: lastChar = $lastChar, nextChar = $nextChar")

    // Если + отсутствует, возвращаем без изменений
    if (!content.contains("+")) {
        println("34.1 В строке отсутствует '+'. Возвращаем без изменений.")
        return content
    }

    // Определяем букву для замены "+"
    val replacement = when {
        lastChar != null && nextChar != null &&
                lastChar in vowels && nextChar in vowels -> "s" // Обе гласные
        lastChar != null && nextChar != null &&
                lastChar !in vowels && nextChar !in vowels -> "i" // Обе согласные
        else -> "" // Если не удаётся определить, оставляем без изменений
    }

    // Заменяем "+" на определённую букву
    val replacedContent = content.replace("+", replacement)
    println("35. Заменённое содержимое с '+': $replacedContent")

    // Проверяем, есть ли символ "*"
    val result = if (replacedContent.contains("*")) {
        replacedContent.replace("*", wordUz)
    } else {
        replacedContent
    }

    println("35.1 Изменённый текст после замены '*': $result")
    return result
}

// Извлекает ячейки из указанного диапазона Excel для дальнейшей обработки.
fun extractCellsFromRange(sheet: Sheet, range: String): List<Cell> {
    println("36. Извлечение ячеек из диапазона: $range")

    // Разделение диапазона на начало и конец
    println("36.1 Разделение диапазона: $range")
    val (start, end) = range.split("-").map {
        println("36.2 Обработка части диапазона: $it")
        it.replace(Regex("[A-Z]"), "").toInt() - 1.also { num ->
            println("36.3 Преобразование в номер строки: $num")
        }
    }
    println("36.4 Диапазон строк: от $start до $end")

    // Определение индекса столбца
    println("36.5 Определение столбца для диапазона: $range")
    val column = range[0] - 'A'
    println("36.6 Индекс столбца: $column")

    // Извлечение ячеек из указанного диапазона
    println("36.7 Извлечение ячеек из строк: $start-$end")
    val cells = (start..end).mapNotNull { rowIndex ->
        println("36.8 Обработка строки: $rowIndex")
        val row = sheet.getRow(rowIndex)
        if (row == null) {
            println("36.9 Строка $rowIndex отсутствует")
            null
        } else {
            val cell = row.getCell(column)
            if (cell == null) {
                println("36.10 Ячейка в строке $rowIndex и столбце $column отсутствует")
                null
            } else {
                println("36.11 Найдена ячейка: строка $rowIndex, столбец $column, значение: ${cell.toString()}")
                cell
            }
        }
    }

    // Итоговый результат
    println("37. Извлечено ${cells.size} ячеек из диапазона $range")
    return cells
}

// Отправляет финальное сообщение пользователю с возможностью выбора следующего действия.
fun sendFinalMessage(chatId: Long, bot: Bot) {
    println("38. Отправка финального сообщения для пользователя $chatId")
    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Поздравляем! Вы прошли все этапы.",
        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Повторить", "repeat"),
            InlineKeyboardButton.CallbackData("Пройти тест", "test")
        )
    )
    println("39. Финальное сообщение отправлено пользователю $chatId")
}
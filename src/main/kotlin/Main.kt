package com.github.kotlintelegrambot.dispatcher

import com.github.kotlintelegrambot.bot
import com.github.kotlintelegrambot.dispatch
import com.github.kotlintelegrambot.dispatcher.command
import com.github.kotlintelegrambot.dispatcher.callbackQuery
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import kotlin.random.Random

val ranges = listOf(
    "A1-A4", "B1-B4", "C1-C7", "D1-D7",
    "A8-A14", "B8-B14", "C8-C14", "D8-D14",
    "A15-A21", "B15-B21", "C15-C21", "D15-D21",
    "A22-A23", "B22-B23", "C22-C28", "D22-D28",
    "A29-A35", "B29-B35", "C29-C35", "D29-D35",
    "A36-A42", "B36-B42", "C36-C42", "D36-D42"
)
val userState = mutableMapOf<Long, Pair<String, Int>>() // <UserId, (currentRange, currentRowIndex)>
var word = "word"

fun main() {
    println("Запуск бота...") // Сообщение о запуске бота
    val tableFile = "Алгоритм 2.8.xlsx" // Имя Excel-файла с данными
    val bot = bot {
        token = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE" // Токен бота

        dispatch {
            command("start") { // Обработка команды /start
                println("Команда /start получена")

                // Создание клавиатуры для выбора слов
                val inlineKeyboard = createWordSelectionKeyboardFromExcel(
                    filePath = tableFile,
                    sheetName = "Существительные"
                )

                // Отправка сообщения с клавиатурой
                bot.sendMessage(
                    chatId = ChatId.fromId(message.chat.id),
                    text = "Выберите слово:",
                    replyMarkup = inlineKeyboard
                )
                println("Сообщение с выбором слов отправлено")
            }

            callbackQuery { // Обработка callback-запросов
                println("Получен callbackQuery: ${callbackQuery.data}") // Логируем полученный callback-запрос

                val chatId = callbackQuery.message?.chat?.id // Получаем идентификатор чата из сообщения
                if (chatId == null) {
                    println("Ошибка: chatId не найден") // Если chatId не найден, выводим сообщение об ошибке
                    return@callbackQuery // Завершаем выполнение этого блока
                }

                val data = callbackQuery.data // Получение данных из callback-запроса
                if (data == null) {
                    println("Ошибка: callbackQuery.data отсутствует") // Если данные отсутствуют, выводим сообщение об ошибке
                    return@callbackQuery // Завершаем выполнение этого блока
                }

                if (data.startsWith("next:")) { // Если запрос содержит "next:", обрабатываем переход к следующему диапазону
                    // Удаляем префикс "next:" из данных
                    val cleanData = data.removePrefix("next:")

                    // Извлекаем текущий диапазон и индекс
                    val (currentRange, currentRowIndex) = userState[chatId] ?: ranges[0] to 0
                    val nextRowIndex = currentRowIndex + 1

                    if (nextRowIndex >= ranges.size) {
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Вы просмотрели все диапазоны!",
                            replyMarkup = null
                        )
                        return@callbackQuery
                    }

                    val nextRange = ranges[nextRowIndex]
                    userState[chatId] = nextRange to nextRowIndex

                    val uzbekWord = cleanData.substringAfter("word:").substringBefore("-").trim()
                    val russianWord = cleanData.substringAfter("-").trim()

                    val messageText = generateMessageFromRange(tableFile, "Примеры гамм для существительны", nextRange)
                        .replace("*", uzbekWord)
                        .replace("#", russianWord)
                        .let {
                            it
                        }

                    bot.sendMessage(
                        chatId = ChatId.fromId(chatId),
                        text = messageText,
                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                            InlineKeyboardButton.CallbackData("Далее", "next:word:$uzbekWord - $russianWord")
                        )
                    )
                }

                if (callbackQuery.data?.startsWith("word:") == true || callbackQuery.data?.startsWith("next:") == false) { // Если запрос содержит "word:" и не содержит "next:", обрабатываем выбор слова
                    val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery // Проверяем наличие chatId
                    val userId = chatId // Идентификатор пользователя совпадает с chatId

                    // Убираем "next:" и "word:", затем делим оставшуюся строку
                    val processedData = data.substringAfter("next:").substringAfter("word:") // Удаляем "next:" и "word:"
                    val uzbekWord = processedData.substringBefore("-").trim() // Берём часть до "-"
                    val russianWord = processedData.substringAfter("-").trim() // Берём часть после "-"

                    println(uzbekWord) // Вывод: uzbekWord
                    println(russianWord) // Вывод: russianWord

                    userState[userId] = ranges[0] to 0 // Устанавливаем начальный диапазон для пользователя

                    val messageText = generateMessageFromRange(tableFile, "Примеры гамм для существительны", ranges[0]) // Генерируем текст из начального диапазона
                        .replace("*", uzbekWord) // Заменяем "*" на узбекское слово
                        .replace("#", russianWord)

                    bot.sendMessage( // Отправляем сформированное сообщение
                        chatId = ChatId.fromId(chatId), // Указываем ID чата
                        text = messageText, // Текст сообщения
                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard( // Добавляем кнопку "Далее"
                            InlineKeyboardButton.CallbackData("Далее", "next: ${callbackQuery.data}")
                        )
                    )
                }
            }


            message { // Обработка обычных сообщений от пользователя
                val userInput = message.text
                if (userInput != null && userInput.contains(" ")) { // Проверка формата ввода
                    val parts = userInput.split(" ", limit = 2)
                    if (parts.size == 2) {
                        val wordUzb = parts[0].trim() // Узбекское слово
                        val wordRus = parts[1].trim() // Русское слово

                        // Добавление слова в таблицу
                        val success = addWordToExcel(
                            filePath = tableFile,
                            sheetName = "Существительные",
                            wordRus = wordRus,
                            wordUzb = wordUzb
                        )

                        // Сообщение об успехе или ошибке
                        if (success) {
                            bot.sendMessage(
                                chatId = ChatId.fromId(message.chat.id),
                                text = "Слово \"$wordRus - $wordUzb\" добавлено.",
                                replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                    InlineKeyboardButton.CallbackData("Выбрать слово", "selectWord"),
                                    InlineKeyboardButton.CallbackData("Добавить еще слово", "addWord")
                                )
                            )
                            println("Слово \"$wordRus - $wordUzb\" добавлено в таблицу")
                        } else {
                            bot.sendMessage(
                                chatId = ChatId.fromId(message.chat.id),
                                text = "Ошибка: не удалось добавить слово. Попробуйте еще раз."
                            )
                        }
                    } else {
                        bot.sendMessage(
                            chatId = ChatId.fromId(message.chat.id),
                            text = "Ошибка: неверный формат. Запишите слово в формате: \"ребенок bola\""
                        )
                    }
                }
            }
        }
    }

    bot.startPolling() // Запуск бота
    println("Бот начал опрос")
}

fun createWordSelectionKeyboardFromExcel(
    filePath: String,
    sheetName: String,
    columnRus: Int = 0, // Номер колонки с русскими словами
    columnUzb: Int = 1  // Номер колонки с узбекскими словами
): InlineKeyboardMarkup {
    // Функция создает клавиатуру для выбора слов из Excel-файла
    println("Чтение слов для клавиатуры из файла: $filePath, лист: $sheetName")
    val file = File(filePath) // Открываем файл
    val workbook = WorkbookFactory.create(file) // Создаем Workbook

    return try {
        val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден") // Проверяем лист
        val rowsCount = sheet.lastRowNum // Получаем общее количество строк
        val randomRows = (0..rowsCount).shuffled().take(10) // Выбираем 10 случайных строк

        val buttons = randomRows.mapNotNull { rowIndex -> // Проходим по выбранным строкам
            val row = sheet.getRow(rowIndex) // Получаем строку
            if (row != null) {
                val wordRus = row.getCell(columnRus)?.toString()?.trim() // Читаем русское слово
                val wordUzb = row.getCell(columnUzb)?.toString()?.trim() // Читаем узбекское слово
                if (!wordRus.isNullOrBlank() && !wordUzb.isNullOrBlank()) {
                    InlineKeyboardButton.CallbackData("$wordRus - $wordUzb", "word:$wordRus - $wordUzb") // Создаем кнопку
                } else {
                    null // Пропускаем строки с пустыми значениями
                }
            } else {
                null
            }
        }.chunked(2) // Разбиваем кнопки на строки по 2 кнопки

        InlineKeyboardMarkup.create(buttons) // Создаем клавиатуру
    } catch (e: Exception) {
        println("Ошибка при создании клавиатуры: ${e.message}") // Логируем ошибки
        InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Ошибка загрузки данных", "error")
        )
    } finally {
        workbook.close() // Закрываем Excel файл
    }
}

fun addWordToExcel(filePath: String, sheetName: String, wordRus: String, wordUzb: String): Boolean {
    // Функция добавляет новое слово в Excel-файл
    println("Добавление слова \"$wordRus - $wordUzb\" в файл: $filePath, лист: $sheetName")
    val file = File(filePath) // Открываем файл
    val workbook = WorkbookFactory.create(file.inputStream()) // Создаем Workbook

    return try {
        val sheet = workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName) // Получаем или создаем лист
        val lastRowNum = if (sheet.physicalNumberOfRows == 0) 0 else sheet.lastRowNum + 1 // Вычисляем индекс новой строки
        val newRow = sheet.createRow(lastRowNum) // Создаем новую строку

        newRow.createCell(0).setCellValue(wordRus) // Добавляем русское слово в первый столбец
        newRow.createCell(1).setCellValue(wordUzb) // Добавляем узбекское слово во второй столбец

        // Сохраняем изменения
        file.outputStream().use { outputStream ->
            workbook.write(outputStream)
        }

        println("Слово \"$wordRus - $wordUzb\" успешно добавлено")
        true // Успешное завершение
    } catch (e: Exception) {
        println("Ошибка при добавлении слова: ${e.message}") // Логируем ошибки
        false // Возвращаем false в случае ошибки
    } finally {
        try {
            workbook.close() // Закрываем Workbook
        } catch (e: Exception) {
            println("Ошибка при закрытии Workbook: ${e.message}") // Логируем ошибки закрытия
        }
    }
}

fun generateMessageFromRange(filePath: String, sheetName: String, range: String): String {
    // Логируем информацию о диапазоне и файле
    println("Чтение диапазона $range из файла $filePath, лист $sheetName")

    val file = File(filePath)
    val workbook = WorkbookFactory.create(file)

    return try {
        val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден")
        val (startCell, endCell) = range.split("-")
        val startRow = startCell.substring(1).toInt() - 1 // Преобразование строки в индекс
        val endRow = endCell.substring(1).toInt() - 1

        val rows = (startRow..endRow).mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex)
            row?.getCell(0)?.toString()?.trim() // Читаем текст из первой ячейки строки
        }

        // Обработка слов с учетом корней и аффиксов
        rows.joinToString("\n") { word ->
            val parts = word.split("*")
            if (parts.size == 2) {
                val root = parts[0].trim()
                val affix = parts[1].trim()

                // Правило добавления соединительного звука
                val connectingSound = determineConnectingSound(root, affix)

                "$root$connectingSound$affix" // Соединяем корень, соединительный звук и аффикс
            } else {
                word // Если формат неправильный, оставляем слово без изменений
            }
        }
    } catch (e: Exception) {
        println("Ошибка при генерации сообщения: ${e.message}")
        "Ошибка: не удалось обработать диапазон $range"
    } finally {
        workbook.close()
    }
}

// Функция определения соединительного звука
fun determineConnectingSound(root: String, affix: String): String {
    // Проверяем последние символы корня и первые символы аффикса
    val lastCharRoot = root.lastOrNull()
    val firstCharAffix = affix.firstOrNull()

    return when {
        lastCharRoot.isVowel() && firstCharAffix.isVowel() -> "s" // Между гласными вставляем "s"
        !lastCharRoot.isVowel() && !firstCharAffix.isVowel() -> "i" // Между согласными вставляем "i"
        else -> "" // Ничего не добавляем, если гласная/согласная и согласная/гласная
    }
}

// Расширение для проверки, является ли символ гласным
fun Char?.isVowel(): Boolean {
    return this != null && this.lowercaseChar() in listOf('a', 'e', 'i', 'o', 'u', 'y')
}




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
    "A1-A7", "B1-B7", "C1-C7", "D1-D7",
    "A8-A14", "B8-B14", "C8-C14", "D8-D14"
)
val userState = mutableMapOf<Long, Pair<String, Int>>() // <UserId, (currentRange, currentRowIndex)>


fun main() {
    println("Запуск бота...") // Сообщение о запуске бота
    val tableFile = "Алгоритм 2.6.xlsx" // Имя Excel-файла с данными
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
                    val (currentRange, currentRowIndex) = userState[chatId] ?: ranges[0] to 0 // Получаем текущий диапазон и индекс строки пользователя
                    val nextRowIndex = currentRowIndex + 1 // Переходим к следующему индексу строки

                    if (nextRowIndex >= ranges.size) { // Проверяем, если индексы закончились
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId), // Отправляем сообщение в чат
                            text = "Вы просмотрели все диапазоны!", // Сообщение пользователю, если диапазоны закончились
                            replyMarkup = null // Убираем кнопки
                        )
                        return@callbackQuery // Завершаем выполнение этого блока
                    }

                    val nextRange = ranges[nextRowIndex] // Получаем следующий диапазон
                    userState[chatId] = nextRange to nextRowIndex // Обновляем состояние пользователя на следующий диапазон

                    val uzbekWord = callbackQuery.data.substringAfter("word:").substringBefore("-").trim() // Извлекаем узбекское слово из данных
                    val russianWord = callbackQuery.data.substringAfter("-").trim() // Извлекаем русское слово из данных

                    val messageText = generateMessageFromRange(tableFile, "Примеры гамм для существительны", nextRange) // Генерируем текст из следующего диапазона
                        .replace("*", uzbekWord) // Заменяем "*" на узбекское слово
                        .replace("#", russianWord) // Заменяем "#" на русское слово
                        .let { "$uzbekWord ($it)" } // Формируем итоговый текст: узбекское слово идет первым, остальной текст внутри скобок

                    bot.sendMessage( // Отправляем сформированное сообщение
                        chatId = ChatId.fromId(chatId), // Указываем ID чата
                        text = messageText, // Текст сообщения
                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard( // Добавляем кнопку "Далее"
                            InlineKeyboardButton.CallbackData("Далее", "next:")
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
                        .replace("#", russianWord) // Заменяем "#" на русское слово
                        .let { "$uzbekWord ($it)" } // Формируем итоговый текст: узбекское слово идет первым, остальной текст внутри скобок

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

fun createWordSelectionKeyboard(): InlineKeyboardMarkup {
    // Функция создает клавиатуру с двумя кнопками
    return InlineKeyboardMarkup.createSingleRowKeyboard(
        InlineKeyboardButton.CallbackData("sabzi", "word:sabzi"), // Кнопка для выбора слова "sabzi"
        InlineKeyboardButton.CallbackData("bola", "word:bola")   // Кнопка для выбора слова "bola"
    )
}

fun createActionKeyboard(selectedWord: String): InlineKeyboardMarkup {
    // Функция создает клавиатуру для действий после выбора слова
    return InlineKeyboardMarkup.createSingleRowKeyboard(
        InlineKeyboardButton.CallbackData("Повторить с $selectedWord", "repeat:$selectedWord"), // Кнопка "Повторить" с выбранным словом
        InlineKeyboardButton.CallbackData("Выбрать новое слово", "selectWord")                // Кнопка для выбора нового слова
    )
}

fun readCellData(filePath: String, sheetName: String, cellReference: String): String {
    // Функция читает данные из указанной ячейки Excel-файла
    println("Попытка чтения данных из файла: $filePath, лист: $sheetName, ячейка: $cellReference")
    val file = File(filePath) // Открываем файл
    val workbook = WorkbookFactory.create(file) // Создаем объект Workbook для работы с Excel
    return try {
        val sheet = workbook.getSheet(sheetName) ?: return "Ошибка: Лист $sheetName не найден" // Получаем лист по имени
        val cellAddress = org.apache.poi.ss.util.CellReference(cellReference) // Преобразуем ссылку на ячейку
        val row = sheet.getRow(cellAddress.row) // Получаем строку по индексу
        if (row == null) {
            println("Ошибка: Строка ${cellAddress.row} отсутствует")
            return "Нет данных" // Если строки нет, возвращаем "Нет данных"
        }
        val cell = row.getCell(cellAddress.col.toInt()) // Получаем ячейку по индексу
        if (cell == null) {
            println("Ошибка: Ячейка ${cellAddress.formatAsString()} пуста")
            return "Нет данных" // Если ячейка пустая, возвращаем "Нет данных"
        }
        cell.toString() // Возвращаем содержимое ячейки в виде строки
    } catch (e: Exception) {
        println("Ошибка чтения Excel: ${e.message}")
        "Ошибка: ${e.message}" // Обрабатываем исключение, если что-то пошло не так
    } finally {
        workbook.close() // Закрываем Excel-файл в любом случае
    }
}

fun selectRandomCellInRange(startCell: String, endCell: String): String {
    // Функция выбирает случайную ячейку в указанном диапазоне
    val startAddress = org.apache.poi.ss.util.CellReference(startCell) // Преобразуем начальную ячейку
    val endAddress = org.apache.poi.ss.util.CellReference(endCell) // Преобразуем конечную ячейку

    val randomRow = (startAddress.row..endAddress.row).random() // Выбираем случайный индекс строки
    val randomCol = (startAddress.col..endAddress.col).random() // Выбираем случайный индекс столбца

    return org.apache.poi.ss.util.CellReference(randomRow, randomCol).formatAsString() // Возвращаем случайную ячейку
}

fun selectRandomCellA(): String {
    // Функция выбирает случайную ячейку из предварительно заданных диапазонов
    val ranges = listOf(
        "A1" to "D1",
        "A8" to "D8",
        "A15" to "D15",
        "A22" to "D22",
        "A29" to "D29",
        "A36" to "D36"
    )
    val (start, end) = ranges.random() // Выбираем случайный диапазон
    val selectedCell = selectRandomCellInRange(start, end) // Выбираем случайную ячейку в этом диапазоне
    println("Выбрана случайная ячейка A: $selectedCell (из диапазона $start-$end)")
    return selectedCell
}

fun selectRandomCellB(cellA: String): String {
    // Функция выбирает ячейку B на основе выбранной ячейки A
    val rangeMap = mapOf(
        "A1-D1" to ("A2" to "A7"),
        "A8-D8" to ("A9" to "A14"),
        "A15-D15" to ("B2" to "B7"),
        "A22-D22" to ("B9" to "B14"),
        "A29-D29" to ("C2" to "C7"),
        "A36-D36" to ("C9" to "C14")
    )

    // Ищем подходящий диапазон для выбранной ячейки A
    val range = rangeMap.entries.find { (key, _) ->
        val (start, end) = key.split("-") // Разделяем диапазон на начальную и конечную ячейки
        cellA in (start..end) // Проверяем, попадает ли cellA в этот диапазон
    }

    return if (range != null) {
        val (start, end) = range.value // Если диапазон найден, выбираем случайную ячейку в нем
        selectRandomCellInRange(start, end)
    } else {
        println("Нет подходящего диапазона для $cellA") // Если диапазон не найден, возвращаем сообщение
        "Нет подходящего диапазона"
    }
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
    println("Чтение диапазона: $range из файла $filePath")
    val file = File(filePath)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName) ?: return "Лист $sheetName не найден"

    val (startCell, endCell) = range.split("-").map { org.apache.poi.ss.util.CellReference(it) }
    val data = StringBuilder()

    for (rowIndex in startCell.row..endCell.row) {
        val row = sheet.getRow(rowIndex)
        val cell = row?.getCell(startCell.col.toInt())
        val cellValue = cell?.toString()?.trim() ?: "Нет данных"
        data.append(cellValue).append("\n")
    }

    return data.toString().trim()
}

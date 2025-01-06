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

// Хранение текущего состояния для каждого пользователя
val userState = mutableMapOf<Long, Pair<String, Int>>() // <UserId, (currentRange, currentRowIndex)>

// Диапазоны ячеек в порядке их обработки
val ranges = listOf(
    "A1-A7", "B1-B7", "C1-C7", "D1-D7",
    "A8-A14", "B8-B14", "C8-C14", "D8-D14",
    "A15-A21", "B15-B21", "C15-C21", "D15-D21",
    "A22-A28", "B22-B28", "C22-C28", "D22-D28",
    "A29-A35", "B29-B35", "C29-C35", "D29-D35",
    "A36-A42", "B36-B42", "C36-C42", "D36-D42"
)

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
                println("Получен callbackQuery: ${callbackQuery.data}")
                val chatId = callbackQuery.message?.chat?.id
                if (chatId == null) {
                    println("Ошибка: chatId не найден") // Проверка наличия chatId
                    return@callbackQuery
                }
                // Получаем текущий диапазон пользователя
                val (currentRange, currentRowIndex) = userState[chatId] ?: ranges[0] to 0
                val nextRowIndex = currentRowIndex + 1

                // Если вышли за пределы диапазонов
                if (nextRowIndex >= ranges.size) {
                    bot.sendMessage(
                        chatId = ChatId.fromId(chatId),
                        text = "Вы просмотрели все диапазоны!",
                        replyMarkup = null
                    )
                    return@callbackQuery
                }

                // Обновляем текущий диапазон
                val data = callbackQuery.data // Получение данных из callback-запроса
                val nextRange = ranges[nextRowIndex]
                userState[chatId] = nextRange to nextRowIndex
                // Формируем сообщение из текущего диапазона
                val messageText = generateMessageFromRange(tableFile, "Примеры гамм для существительны", nextRange)
                bot.sendMessage(
                    chatId = ChatId.fromId(chatId),
                    text = messageText,
                    replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                        InlineKeyboardButton.CallbackData("Далее", "next:")
                    )
                )

                if (data == null) {
                    println("Ошибка: callbackQuery.data отсутствует") // Проверка данных
                }

                when {
                    data.startsWith("word:") -> { val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery
                        val userId = chatId

                        // Инициализируем пользователя с начальным диапазоном
                        userState[userId] = ranges[0] to 0

                        // Генерируем первое сообщение из диапазона A1-A7
                        val messageText = generateMessageFromRange(tableFile, "Примеры гамм для существительны", ranges[0])
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = messageText,
                            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                InlineKeyboardButton.CallbackData("Далее", "next:")
                            )
                        )
                    }

                    data.startsWith("repeat:") -> { // Обработка нажатия кнопки "Повторить"
                        println("Пользователь запросил повтор: $data")
                        val buttonText = data.removePrefix("repeat:").trim()
                        val parts = buttonText.split(" - ")
                        if (parts.size != 2) {
                            println("Ошибка: Неверный формат текста кнопки: $buttonText")
                            return@callbackQuery
                        }

                        // Извлечение выбранных слов
                        val wordRus = parts[0]
                        val wordUzb = parts[1]

                        var dataA: String
                        var dataB: String
                        var resultA: String
                        var resultB: String

                        // Повторный цикл выбора ячеек
                        do {
                            resultA = selectRandomCellA()
                            println("Выбрана ячейка A: $resultA")
                            dataA = readCellData(tableFile, "Примеры гамм для существительны", resultA)
                            println("Прочитаны данные A: $dataA")

                            resultB = selectRandomCellB(resultA)
                            println("Выбрана ячейка B: $resultB")
                            dataB = readCellData(tableFile, "Примеры гамм для существительны", resultB)

                            if (dataB.isBlank() || dataB == "Нет данных") {
                                println("Ячейка B ($resultB) пуста. Повторный выбор ячеек.")
                            }
                        } while (dataB.isBlank() || dataB == "Нет данных")

                        println("Прочитаны данные B: $dataB")

                        // Формирование текста сообщения
                        val messageText = """
        $dataA
        ${dataB.replace("*", wordUzb).replace("#", wordRus)}
    """.trimIndent()

                        // Отправка сообщения
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = messageText,
                            replyMarkup = createActionKeyboard(buttonText)
                        )
                        println("Сообщение отправлено пользователю")
                    }

                    data == "selectWord" -> { // Обработка нажатия кнопки "Выбрать новое слово"
                        println("Пользователь запросил выбор нового слова")

                        // Создание новой клавиатуры для выбора слова
                        val inlineKeyboard = createWordSelectionKeyboardFromExcel(
                            filePath = tableFile,
                            sheetName = "Существительные"
                        )

                        // Отправка клавиатуры
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Выберите слово:",
                            replyMarkup = inlineKeyboard
                        )
                        println("Сообщение с клавиатурой выбора отправлено")
                    }

                    data == "addWord" -> { // Обработка добавления нового слова
                        println("Пользователь запросил добавление слова")
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Чтобы добавить слово, запишите его в следующем формате: \"ребенок bola\""
                        )
                        println("Инструкция отправлена пользователю")
                    }

                    else -> {
                        println("Неизвестная команда callbackQuery: $data") // Если данные не совпадают с известными командами
                    }
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

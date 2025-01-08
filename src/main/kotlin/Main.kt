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

fun main() {
    println("Запуск бота...") // Сообщение о запуске бота
    val tableFile = "Алгоритм 2.9.xlsx" // Имя Excel-файла с данными
    val bot = bot {
        token = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE" // Токен бота

        dispatch {
            command("start") { // Обработка команды /start
                println("1. Команда /start получена")

                // 2. Создание клавиатуры для выбора слов
                val inlineKeyboard = createWordSelectionKeyboardFromExcel(
                    filePath = tableFile,
                    sheetName = "Существительные"
                )
                println("2. Клавиатура для выбора слов создана")

                // 3. Отправка сообщения с клавиатурой
                bot.sendMessage(
                    chatId = ChatId.fromId(message.chat.id),
                    text = "Выберите слово:",
                    replyMarkup = inlineKeyboard
                )
                println("3. Сообщение с выбором слов отправлено")
            }

            callbackQuery { // Обработка callback-запросов
                println("4. Получен callbackQuery: ${callbackQuery.data}") // Логируем полученный callback-запрос

                val chatId = callbackQuery.message?.chat?.id // Получаем идентификатор чата из сообщения
                if (chatId == null) {
                    println("4.1. Ошибка: chatId не найден") // Если chatId не найден, выводим сообщение об ошибке
                    return@callbackQuery // Завершаем выполнение этого блока
                }

                val data = callbackQuery.data // Получение данных из callback-запроса
                if (data == null) {
                    println("4.2. Ошибка: callbackQuery.data отсутствует") // Если данные отсутствуют, выводим сообщение об ошибке
                    return@callbackQuery // Завершаем выполнение этого блока
                }

                println("5. Данные из callbackQuery: $data")

                val uzbekWord = data.substringAfter("word:").substringBefore("-").trim()
                println("6. Извлечён узбекское слово: $uzbekWord")

                if (data.startsWith("next:")) { // Если запрос содержит "next:", обрабатываем переход к следующему диапазону
                    println("7. Обработка следующего диапазона")

                    val cleanData = data.removePrefix("next:")
                    println("7.1. Данные после удаления 'next:': $cleanData")

                    val (currentRange, currentRowIndex) = userState[chatId] ?: ranges[0] to 0
                    val nextRowIndex = currentRowIndex + 1
                    println("7.2. Текущий диапазон: $currentRange, текущий индекс строки: $currentRowIndex")

                    if (nextRowIndex >= ranges.size) {
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Вы просмотрели все диапазоны!",
                            replyMarkup = null
                        )
                        println("7.3. Все диапазоны просмотрены")
                        return@callbackQuery
                    }

                    val nextRange = ranges[nextRowIndex]
                    userState[chatId] = nextRange to nextRowIndex
                    println("7.4. Следующий диапазон: $nextRange, обновлён индекс строки: $nextRowIndex")

                    val uzbekWord = cleanData.substringAfter("word:").substringBefore("-").trim()
                    val russianWord = cleanData.substringAfter("-").trim()
                    println("7.5. Узбекское слово: $uzbekWord, Русское слово: $russianWord")

                    val messageText =
                        generateMessageFromRange(tableFile, "Примеры гамм для существительны", nextRange, uzbekWord)
                            .replace("*", uzbekWord)
                            .replace("#", russianWord)
                    println("7.6. Сформированный текст сообщения: $messageText")

                    bot.sendMessage(
                        chatId = ChatId.fromId(chatId),
                        text = messageText,
                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                            InlineKeyboardButton.CallbackData("Далее", "next:word:$uzbekWord - $russianWord")
                        )
                    )
                    println("7.7. Сообщение отправлено")
                }

                if (callbackQuery.data?.startsWith("word:") == true || callbackQuery.data?.startsWith("next:") == false) {
                    println("8. Обработка выбора слова")

                    val processedData =
                        data.substringAfter("next:").substringAfter("word:") // Удаляем "next:" и "word:"
                    val uzbekWord = processedData.substringBefore("-").trim() // Берём часть до "-"
                    val russianWord = processedData.substringAfter("-").trim() // Берём часть после "-"
                    println("8.1. Узбекское слово: $uzbekWord, Русское слово: $russianWord")

                    userState[chatId] = ranges[0] to 0 // Устанавливаем начальный диапазон для пользователя
                    println("8.2. Установлен начальный диапазон для пользователя")

                    val messageText =
                        generateMessageFromRange(tableFile, "Примеры гамм для существительны", ranges[0], uzbekWord)
                            .replace("*", uzbekWord)
                            .replace("#", russianWord)
                    println("8.3. Сформированный текст сообщения: $messageText")

                    bot.sendMessage(
                        chatId = ChatId.fromId(chatId),
                        text = messageText,
                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                            InlineKeyboardButton.CallbackData("Далее", "next: ${callbackQuery.data}")
                        )
                    )
                    println("8.4. Сообщение отправлено")
                }
            }

            message { // Обработка обычных сообщений от пользователя
                println("9. Получено сообщение от пользователя")
                val userInput = message.text
                println("9.1. Текст сообщения: $userInput")

                if (userInput != null && userInput.contains(" ")) { // Проверка формата ввода
                    println("9.2. Проверка формата ввода прошла успешно")

                    val parts = userInput.split(" ", limit = 2)
                    println("9.3. Сообщение разделено на части: $parts")

                    if (parts.size == 2) {
                        val wordUzb = parts[0].trim() // Узбекское слово
                        val wordRus = parts[1].trim() // Русское слово
                        println("9.4. Узбекское слово: $wordUzb, Русское слово: $wordRus")

                        // 10. Добавление слова в таблицу
                        val success = addWordToExcel(
                            filePath = tableFile,
                            sheetName = "Существительные",
                            wordRus = wordRus,
                            wordUzb = wordUzb
                        )
                        println("10. Попытка добавить слово в таблицу: $success")

                        // 11. Сообщение об успехе или ошибке
                        if (success) {
                            bot.sendMessage(
                                chatId = ChatId.fromId(message.chat.id),
                                text = "Слово \"$wordRus - $wordUzb\" добавлено.",
                                replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                    InlineKeyboardButton.CallbackData("Выбрать слово", "selectWord"),
                                    InlineKeyboardButton.CallbackData("Добавить еще слово", "addWord")
                                )
                            )
                            println("11. Слово \"$wordRus - $wordUzb\" успешно добавлено в таблицу")
                        } else {
                            bot.sendMessage(
                                chatId = ChatId.fromId(message.chat.id),
                                text = "Ошибка: не удалось добавить слово. Попробуйте еще раз."
                            )
                            println("11. Ошибка: не удалось добавить слово \"$wordRus - $wordUzb\"")
                        }
                    } else {
                        println("9.5. Некорректный формат ввода")
                        bot.sendMessage(
                            chatId = ChatId.fromId(message.chat.id),
                            text = "Ошибка: неверный формат. Запишите слово в формате: \"ребенок bola\""
                        )
                    }
                } else {
                    println("9.6. Сообщение не содержит пробела или пустое")
                }
            }
        }
    }
    println("12. Бот начал опрос")
    bot.startPolling() // Запуск бота
}

fun createWordSelectionKeyboardFromExcel( // Создаёт клавиатуру для выбора слов на основе данных из Excel-файла.
    filePath: String,
    sheetName: String,
    columnRus: Int = 0, // Номер колонки с русскими словами
    columnUzb: Int = 1  // Номер колонки с узбекскими словами
): InlineKeyboardMarkup {
    println("13. Чтение слов для клавиатуры из файла: $filePath, лист: $sheetName")
    val file = File(filePath) // Открываем файл
    println("13.1. Файл успешно открыт: ${file.absolutePath}")
    val workbook = WorkbookFactory.create(file) // Создаем Workbook из Excel-файла
    println("13.2. Workbook успешно создан")

    return try {
        val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден") // Проверяем лист
        println("13.3. Лист $sheetName найден, общее количество строк: ${sheet.lastRowNum}")

        val rowsCount = sheet.lastRowNum // Получаем общее количество строк
        val randomRows = (0..rowsCount).shuffled().take(10) // Выбираем 10 случайных строк
        println("13.4. Выбраны случайные строки: $randomRows")

        val buttons = randomRows.mapNotNull { rowIndex -> // Проходим по выбранным строкам
            println("13.5. Обработка строки с индексом $rowIndex")
            val row = sheet.getRow(rowIndex) // Получаем строку
            if (row != null) {
                val wordRus = row.getCell(columnRus)?.toString()?.trim() // Читаем русское слово
                val wordUzb = row.getCell(columnUzb)?.toString()?.trim() // Читаем узбекское слово
                println("13.6. Русское слово: $wordRus, Узбекское слово: $wordUzb")

                if (!wordRus.isNullOrBlank() && !wordUzb.isNullOrBlank()) {
                    InlineKeyboardButton.CallbackData("$wordRus - $wordUzb", "word:$wordRus - $wordUzb") // Создаем кнопку
                } else {
                    println("13.7. Пропущена строка с пустыми значениями")
                    null // Пропускаем строки с пустыми значениями
                }
            } else {
                println("13.8. Строка $rowIndex отсутствует")
                null
            }
        }.chunked(2) // Разбиваем кнопки на строки по 2 кнопки
        println("13.9. Кнопки сформированы")

        InlineKeyboardMarkup.create(buttons) // Создаем клавиатуру
    } catch (e: Exception) {
        println("13.10. Ошибка при создании клавиатуры: ${e.message}") // Логируем ошибки
        InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Ошибка загрузки данных", "error")
        )
    } finally {
        workbook.close() // Закрываем Excel файл
        println("13.11. Workbook закрыт")
    }
}

fun addWordToExcel(filePath: String, sheetName: String, wordRus: String, wordUzb: String): Boolean { // Добавляет новое слово в указанный Excel-файл.
    println("14. Добавление слова \"$wordRus - $wordUzb\" в файл: $filePath, лист: $sheetName")
    val file = File(filePath) // Открываем файл
    println("14.1. Файл успешно открыт: ${file.absolutePath}")
    val workbook = WorkbookFactory.create(file.inputStream()) // Создаем Workbook
    println("14.2. Workbook успешно создан")

    return try {
        val sheet = workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName) // Получаем или создаем лист
        println("14.3. Лист $sheetName успешно получен/создан")

        val lastRowNum = if (sheet.physicalNumberOfRows == 0) 0 else sheet.lastRowNum + 1 // Вычисляем индекс новой строки
        println("14.4. Индекс новой строки: $lastRowNum")

        val newRow = sheet.createRow(lastRowNum) // Создаем новую строку
        newRow.createCell(0).setCellValue(wordRus) // Добавляем русское слово в первый столбец
        newRow.createCell(1).setCellValue(wordUzb) // Добавляем узбекское слово во второй столбец
        println("14.5. Слово \"$wordRus - $wordUzb\" добавлено в строку $lastRowNum")

        // Сохраняем изменения
        file.outputStream().use { outputStream ->
            workbook.write(outputStream)
        }
        println("14.6. Изменения сохранены в файл")

        true // Успешное завершение
    } catch (e: Exception) {
        println("14.7. Ошибка при добавлении слова: ${e.message}") // Логируем ошибки
        false // Возвращаем false в случае ошибки
    } finally {
        try {
            workbook.close() // Закрываем Workbook
            println("14.8. Workbook закрыт")
        } catch (e: Exception) {
            println("14.9. Ошибка при закрытии Workbook: ${e.message}") // Логируем ошибки закрытия
        }
    }
}

fun generateMessageFromRange(filePath: String, sheetName: String, range: String, uzbekWord: String): String {
    // Создаёт сообщение на основе диапазона данных из Excel-файла.

    // 15. Логируем информацию о диапазоне и файле
    println("15. Чтение диапазона $range из файла $filePath, лист $sheetName")

    val file = File(filePath) // Открываем файл по указанному пути
    println("16. Файл для чтения: ${file.absolutePath}")

    val workbook = WorkbookFactory.create(file) // Создаем Workbook из Excel-файла
    println("17. Excel-файл успешно открыт")

    return try {
        // 18. Получаем указанный лист из Excel-файла
        val sheet = workbook.getSheet(sheetName) ?: throw IllegalArgumentException("Лист $sheetName не найден")
        println("18. Лист Excel найден: $sheetName")

        // 19. Разбиваем диапазон на начало и конец (например, "A1-A4" на "A1" и "A4")
        val (startCell, endCell) = range.split("-")
        println("19. Диапазон разбит: начало $startCell, конец $endCell")

        // 20. Определяем номер столбца по букве
        val startColumnLetter = startCell.substring(0, 1) // Буква начального столбца
        val columnIndex = startColumnLetter[0] - 'A' // Преобразование буквы в индекс столбца (A=0, B=1, ...)
        println("20. Определён столбец: $startColumnLetter, индекс: $columnIndex")

        // 21. Преобразуем строковые обозначения в индексы строк
        val startRow = startCell.substring(1).toInt() - 1 // Индекс строки начала диапазона
        val endRow = endCell.substring(1).toInt() - 1 // Индекс строки конца диапазона
        println("21. Индексы строк: начало $startRow, конец $endRow")

        // 22. Читаем строки из диапазона и извлекаем текст из указанного столбца каждой строки
        val rows = (startRow..endRow).mapNotNull { rowIndex ->
            val row = sheet.getRow(rowIndex) // Получаем строку по индексу
            val cellValue = row?.getCell(columnIndex)?.toString()?.trim() // Читаем и очищаем текст из указанного столбца
            val cellAddress = "${startColumnLetter}${rowIndex + 1}" // Формируем адрес ячейки
            println("22.1. Чтение строки с индексом $rowIndex ($cellAddress): $cellValue") // Логирование

            if (cellValue != null) cellAddress to cellValue else null // Возвращаем пару: адрес ячейки и её содержимое
        }
        println("23. Содержимое строк: $rows")

        // 24. Обрабатываем каждую строку, добавляя адрес ячейки
        rows.joinToString("\n") { (cellAddress, word) -> // Соединяем обработанные строки в одно сообщение
            println("24. Обработка строки ($cellAddress): $word")
            if (word.contains("*")) { // Если в строке есть звёздочка
                val prefix = word.substringBefore("*").trim() // Текст перед звёздочкой
                val suffix = word.substringAfter("*").trim() // Текст после звёздочки
                println("24.1. Префикс: $prefix, Суффикс: $suffix")

                if (startColumnLetter == "C") {
                    // Только для столбца C добавляем соединительный звук
                    val connectingSound = determineConnectingSound(prefix, uzbekWord)
                    println("24.2. Соединительный звук: $connectingSound")
                    "$cellAddress: $prefix $connectingSound$uzbekWord$suffix" // Сохраняем пробел между частями
                } else {
                    // Для остальных столбцов просто заменяем *
                    "$cellAddress: $prefix $uzbekWord $suffix" // Сохраняем пробел между частями
                }
            } else {
                println("24.3. Звёздочка отсутствует, возвращается оригинал")
                "$cellAddress: $word" // Если звёздочки нет, добавляем адрес ячейки к оригиналу
            }
        }
    } catch (e: Exception) {
        // 25. Обрабатываем ошибки при чтении или обработке Excel-файла
        println("25. Ошибка при генерации сообщения: ${e.message}")
        "Ошибка: не удалось обработать диапазон $range" // Возвращаем сообщение об ошибке
    } finally {
        workbook.close() // 26. Закрываем Workbook, чтобы освободить ресурсы
        println("26. Excel-файл закрыт")
    }
}


fun determineConnectingSound(prefix: String, uzbekWord: String): String {
    // Определяет, какой соединительный звук ("s" или "i") нужно вставить между корнем и узбекским словом.

    // 26. Проверяем последний символ префикса
    val lastCharPrefix = prefix.lastOrNull() // Получаем последний символ префикса
    println("26. Последний символ префикса: $lastCharPrefix")

    // 27. Проверяем первый символ узбекского слова
    val firstCharUzbekWord = uzbekWord.firstOrNull() // Получаем первый символ узбекского слова
    println("27. Первый символ узбекского слова: $firstCharUzbekWord")

    // 28. Выбираем соединительный звук в зависимости от типа символов
    return when {
        lastCharPrefix.isVowel() && firstCharUzbekWord.isVowel() -> {
            println("28. Между гласными добавляем 's'")
            "s" // Если оба символа гласные, вставляем "s"
        }
        !lastCharPrefix.isVowel() && !firstCharUzbekWord.isVowel() -> {
            println("28. Между согласными добавляем 'i'")
            "i" // Если оба символа согласные, вставляем "i"
        }
        else -> {
            println("28. Соединительный звук не добавляется")
            "" // В остальных случаях ничего не добавляем
        }
    }
}

// Расширение для проверки, является ли символ гласным
fun Char?.isVowel(): Boolean {
    return this != null && this.lowercaseChar() in listOf('a', 'e', 'i', 'o', 'u', 'y')
}
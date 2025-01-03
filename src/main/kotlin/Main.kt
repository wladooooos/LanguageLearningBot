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

fun main() {
    println("Запуск бота...")

    val bot = bot {
        token = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE"

        dispatch {
            command("start") {
                println("Команда /start получена")
                val inlineKeyboard = createWordSelectionKeyboard()
                bot.sendMessage(
                    chatId = ChatId.fromId(message.chat.id),
                    text = "Выберите слово:",
                    replyMarkup = inlineKeyboard
                )
                println("Сообщение с выбором слов отправлено")
            }

            callbackQuery {
                println("Получен callbackQuery: ${callbackQuery.data}")
                val chatId = callbackQuery.message?.chat?.id
                if (chatId == null) {
                    println("Ошибка: chatId не найден")
                    return@callbackQuery
                }

                val data = callbackQuery.data
                if (data == null) {
                    println("Ошибка: callbackQuery.data отсутствует")
                    return@callbackQuery
                }

                when {
                    data.startsWith("word:") -> {
                        println("Пользователь выбрал слово: $data")
                        val selectedWord = data.removePrefix("word:")

                        var dataA: String
                        var dataB: String
                        var resultA: String
                        var resultB: String

                        do {
                            resultA = selectRandomCellA()
                            println("Выбрана ячейка A: $resultA")
                            dataA = readCellData("Алгоритм 2.3.xlsx", "Примеры гамм для существительны", resultA)
                            println("Прочитаны данные A: $dataA")

                            resultB = selectRandomCellB(resultA)
                            println("Выбрана ячейка B: $resultB")
                            dataB = readCellData("Алгоритм 2.3.xlsx", "Примеры гамм для существительны", resultB)

                            if (dataB.isBlank() || dataB == "Нет данных") {
                                println("Ячейка B ($resultB) пуста. Повторный выбор ячеек.")
                            }
                        } while (dataB.isBlank() || dataB == "Нет данных")

                        println("Прочитаны данные B: $dataB")

                        val messageText = """
        $dataA
        ${dataB.replace("*", selectedWord)}
    """.trimIndent()

                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = messageText,
                            replyMarkup = createActionKeyboard(selectedWord)
                        )
                        println("Сообщение отправлено пользователю")
                    }

                    data.startsWith("repeat:") -> {
                        println("Пользователь запросил повтор: $data")
                        val selectedWord = data.removePrefix("repeat:")

                        var dataA: String
                        var dataB: String
                        var resultA: String
                        var resultB: String

                        do {
                            resultA = selectRandomCellA()
                            println("Выбрана ячейка A: $resultA")
                            dataA = readCellData("Алгоритм 2.3.xlsx", "Примеры гамм для существительны", resultA)
                            println("Прочитаны данные A: $dataA")

                            resultB = selectRandomCellB(resultA)
                            println("Выбрана ячейка B: $resultB")
                            dataB = readCellData("Алгоритм 2.3.xlsx", "Примеры гамм для существительны", resultB)

                            if (dataB.isBlank() || dataB == "Нет данных") {
                                println("Ячейка B ($resultB) пуста. Повторный выбор ячеек.")
                            }
                        } while (dataB.isBlank() || dataB == "Нет данных")

                        println("Прочитаны данные B: $dataB")

                        val messageText = """
        Информация: $dataA
        В: ${dataB.replace("*", selectedWord)}
    """.trimIndent()

                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = messageText,
                            replyMarkup = createActionKeyboard(selectedWord)
                        )
                        println("Сообщение отправлено пользователю")
                    }

                    data == "selectWord" -> {
                        println("Пользователь запросил выбор нового слова")
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = "Выберите слово:",
                            replyMarkup = createWordSelectionKeyboard()
                        )
                        println("Сообщение с клавиатурой выбора отправлено")
                    }
                    else -> {
                        println("Неизвестная команда callbackQuery: $data")
                    }
                }
            }
        }
    }

    bot.startPolling()
    println("Бот начал опрос")
}

fun createWordSelectionKeyboard(): InlineKeyboardMarkup {
    return InlineKeyboardMarkup.createSingleRowKeyboard(
        InlineKeyboardButton.CallbackData("sabzi", "word:sabzi"),
        InlineKeyboardButton.CallbackData("bola", "word:bola")
    )
}

fun createActionKeyboard(selectedWord: String): InlineKeyboardMarkup {
    return InlineKeyboardMarkup.createSingleRowKeyboard(
        InlineKeyboardButton.CallbackData("Повторить с $selectedWord", "repeat:$selectedWord"),
        InlineKeyboardButton.CallbackData("Выбрать новое слово", "selectWord")
    )
}

fun readCellData(filePath: String, sheetName: String, cellReference: String): String {
    println("Попытка чтения данных из файла: $filePath, лист: $sheetName, ячейка: $cellReference")
    val file = File(filePath)
    val workbook = WorkbookFactory.create(file)
    return try {
        val sheet = workbook.getSheet(sheetName) ?: return "Ошибка: Лист $sheetName не найден"
        val cellAddress = org.apache.poi.ss.util.CellReference(cellReference)
        val row = sheet.getRow(cellAddress.row)
        if (row == null) {
            println("Ошибка: Строка ${cellAddress.row} отсутствует")
            return "Нет данных"
        }
        val cell = row.getCell(cellAddress.col.toInt())
        if (cell == null) {
            println("Ошибка: Ячейка ${cellAddress.formatAsString()} пуста")
            return "Нет данных"
        }
        cell.toString()
    } catch (e: Exception) {
        println("Ошибка чтения Excel: ${e.message}")
        "Ошибка: ${e.message}"
    } finally {
        workbook.close() // Закрываем файл в любом случае
    }
}


fun selectRandomCellInRange(startCell: String, endCell: String): String {
    val startAddress = org.apache.poi.ss.util.CellReference(startCell)
    val endAddress = org.apache.poi.ss.util.CellReference(endCell)

    val randomRow = (startAddress.row..endAddress.row).random()
    val randomCol = (startAddress.col..endAddress.col).random()

    return org.apache.poi.ss.util.CellReference(randomRow, randomCol).formatAsString()
}

fun selectRandomCellA(): String {
    val ranges = listOf(
        "A1" to "D1",
        "A8" to "D8",
        "A15" to "D15",
        "A22" to "D22",
        "A29" to "D29",
        "A36" to "D36"
    )
    val (start, end) = ranges.random()
    val selectedCell = selectRandomCellInRange(start, end)
    println("Выбрана случайная ячейка A: $selectedCell (из диапазона $start-$end)")
    return selectedCell
}


fun selectRandomCellB(cellA: String): String {
    val rangeMap = mapOf(
        "A1-D1" to ("A2" to "A7"),
        "A8-D8" to ("A9" to "A14"),
        "A15-D15" to ("B2" to "B7"),
        "A22-D22" to ("B9" to "B14"),
        "A29-D29" to ("C2" to "C7"),
        "A36-D36" to ("C9" to "C14")
    )

    val range = rangeMap.entries.find { (key, _) ->
        val (start, end) = key.split("-")
        cellA in (start..end)
    }

    return if (range != null) {
        val (start, end) = range.value
        selectRandomCellInRange(start, end)
    } else {
        println("Нет подходящего диапазона для $cellA")
        "Нет подходящего диапазона"
    }
}


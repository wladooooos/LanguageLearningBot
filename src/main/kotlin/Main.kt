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
var wordUz: String? = "bola"
var wordRus: String? = "ребенок"
val userReplacements = mutableMapOf<Long, Map<Int, String>>() // Хранение замен для чисел (1-9) для каждого пользователя
var sheetColumnPairs = mutableMapOf<Long, Map<String, String>>() // Глобальная переменная для хранения пар лист/столбец (ключ и значение — строки)

val tableFile = "Table.xlsx"

fun main() {
    println("ЖЖЖ main // Основная функция запуска бота")
    val bot = bot {
        token = "7856005284:AAFVvPnRadWhaotjUZOmFyDFgUHhZ0iGsCo"
        println("Ж1 Инициализация бота с токеном")

        dispatch {
            command("start") {
                val chatId = message.chat.id
                println("Ж2 Команда /start от пользователя: chatId = $chatId")

                // Сброс состояния
                userStates.remove(chatId)
                userPadezh.remove(chatId)
                userBlocks[chatId] = 1
                userBlockCompleted.remove(chatId)
                userColumnOrder.remove(chatId)
                println("Ж3 Сброс состояния для пользователя: chatId = $chatId")

                sendWelcomeMessage(chatId, bot) // Приветственное сообщение
                sendStartMenu(chatId, bot) // Отправляем стартовое меню
            }


            callbackQuery {
                val chatId = callbackQuery.message?.chat?.id ?: return@callbackQuery
                val data = callbackQuery.data ?: return@callbackQuery
                println("Ж6 Получен callback от пользователя: chatId = $chatId, data = $data")

                when {
                    data == "main_menu" -> {
                        println("🔄 Возврат в начальное меню для chatId = $chatId")

                        // Сбрасываем все состояния пользователя
                        userStates.remove(chatId)
                        userPadezh.remove(chatId)
                        userWords.remove(chatId)
                        userBlocks.remove(chatId)
                        userBlockCompleted.remove(chatId)
                        userColumnOrder.remove(chatId)

                        // Отправляем стартовое меню
                        sendStartMenu(chatId, bot)
                    }
                    data.startsWith("Padezh:") -> {
                        val selectedPadezh = data.removePrefix("Padezh:")
                        println("Ж7 Выбранный падеж: chatId = $chatId, selectedPadezh = $selectedPadezh")
                        userPadezh[chatId] = selectedPadezh
                        userStates[chatId] = 0
                        userColumnOrder.remove(chatId)
                        bot.sendMessage(
                            chatId = ChatId.fromId(chatId),
                            text = """Падеж: $selectedPadezh
                                |Слово: $wordUz - $wordRus
                            """.trimMargin()
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
                    data.startsWith("next_adjective:") -> {
                        println("Ж18 Обработка 'Далее' для блока прилагательных: data = $data")

                        val currentState = userStates[chatId] ?: 0
                        userStates[chatId] = currentState + 1
                        println("Ж19 Обновлено состояние пользователя: chatId=$chatId, newState=${userStates[chatId]}")

                        handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                    }
                    data.startsWith("block:") -> {
                        val blockType = data.removePrefix("block:")
                        when (blockType) {
                            "1" -> {
                                userBlocks[chatId] = 1
                                println("Пользователь $chatId выбрал блок 1")
                                handleBlock1(chatId, bot, tableFile, wordUz, wordRus)
                            }
                            "2" -> {
                                initializeUserBlockStates(chatId, tableFile)
                                val (block1Completed, _, _) = userBlockCompleted[chatId] ?: Triple(false, false, false)

                                if (block1Completed) {
                                    userBlocks[chatId] = 2
                                    println("✅ Пользователь $chatId перешел во 2-й блок")
                                    handleBlock2(chatId, bot, tableFile, wordUz, wordRus)
                                } else {
                                    val messageText = "Вы не завершили Блок 1.\nПройдите его перед переходом ко 2-му блоку."
                                    bot.sendMessage(
                                        chatId = ChatId.fromId(chatId),
                                        text = messageText,
                                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                            InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
                                        )
                                    )
                                    println("⚠️ Пользователь $chatId не завершил Блок 1. Доступ к Блоку 2 закрыт.")
                                }
                            }

                            "3" -> {
                                initializeUserBlockStates(chatId, tableFile)
                                val (block1Completed, block2Completed, _) = userBlockCompleted[chatId] ?: Triple(false, false, false)

                                if (block1Completed && block2Completed) {
                                    userBlocks[chatId] = 3
                                    println("✅ Пользователь $chatId перешел в 3-й блок")
                                    handleBlock3(chatId, bot, tableFile, wordUz, wordRus)
                                } else {
                                    val notCompletedBlocks = mutableListOf<String>()
                                    if (!block1Completed) notCompletedBlocks.add("Блок 1")
                                    if (!block2Completed) notCompletedBlocks.add("Блок 2")

                                    val messageText = "Вы не завершили следующие блоки:\n" +
                                            notCompletedBlocks.joinToString("\n") + "\nПройдите их перед переходом к 3-му блоку."

                                    bot.sendMessage(
                                        chatId = ChatId.fromId(chatId),
                                        text = messageText,
                                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                            InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
                                        )
                                    )
                                    println("⚠️ Пользователь $chatId не завершил блоки: $notCompletedBlocks. Доступ к Блоку 3 закрыт.")
                                }
                            }

                            "test" -> {
                                initializeUserBlockStates(chatId, tableFile)
                                val (block1Completed, block2Completed, block3Completed) = userBlockCompleted[chatId] ?: Triple(false, false, false)

                                if (block1Completed && block2Completed && block3Completed) {
                                    userBlocks[chatId] = 4
                                    println("✅ Пользователь $chatId перешел в тестовый блок")
                                    checkBlocksBeforeTest(chatId, bot, tableFile)
                                } else {
                                    val notCompletedBlocks = mutableListOf<String>()
                                    if (!block1Completed) notCompletedBlocks.add("Блок 1")
                                    if (!block2Completed) notCompletedBlocks.add("Блок 2")
                                    if (!block3Completed) notCompletedBlocks.add("Блок 3")

                                    val messageText = "Вы не завершили следующие блоки:\n" +
                                            notCompletedBlocks.joinToString("\n") + "\nПройдите их перед тестом."

                                    bot.sendMessage(
                                        chatId = ChatId.fromId(chatId),
                                        text = messageText,
                                        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                            InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
                                        )
                                    )
                                    println("⚠️ Пользователь $chatId не завершил блоки: $notCompletedBlocks. Доступ к тесту закрыт.")
                                }
                            }
                            "adjective1" -> {
                                userBlocks[chatId] = 5
                                userStates.remove(chatId)
                                userPadezh.remove(chatId)
                                userColumnOrder.remove(chatId)

                                println("Пользователь $chatId выбрал блок прилагательных 1")
                                handleBlockAdjective1(chatId, bot)
                            }
                            "adjective2" -> {
                                userBlocks[chatId] = 6
                                userStates.remove(chatId)
                                userPadezh.remove(chatId)
                                userColumnOrder.remove(chatId)
                                println("Пользователь $chatId выбрал блок прилагательных 2")
                                handleBlockAdjective2(chatId, bot)
                            }
                            else -> {
                                bot.sendMessage(
                                    chatId = ChatId.fromId(chatId),
                                    text = "Неизвестный блок: $blockType"
                                )
                                println("Ж28 Неизвестный блок: $blockType")
                            }
                        }
                    }
                    data == "change_word" -> {
                        println("Ж20 Выбор нового слова для пользователя: chatId = $chatId")
                        userStates[chatId] = 0
                        userWords.remove(chatId)
                        wordUz = null
                        wordRus = null
                        userColumnOrder.remove(chatId)
                        sendWordMessage(chatId, bot, tableFile)
                        println("Ж21111 Отправлена клавиатура для выбора слов: chatId = $chatId")
                    }
                    data == "change_words_adjective1" -> {
                        println("🔄 Перегенерация набора слов для прилагательных 1: chatId = $chatId")
                        userReplacements.remove(chatId)
                        sheetColumnPairs.remove(chatId)
                        userStates.remove(chatId)
                        initializeSheetColumnPairsFromFile(chatId)  // Генерация новых пар слов
                        handleBlockAdjective1(chatId, bot) // Перезапуск блока 1
                    }
                    data == "change_words_adjective2" -> {
                        println("🔄 Перегенерация набора слов для прилагательных 2: chatId = $chatId")
                        userReplacements.remove(chatId)
                        sheetColumnPairs.remove(chatId)
                        userStates.remove(chatId)
                        initializeSheetColumnPairsFromFile(chatId)  // Генерация новых пар слов
                        handleBlockAdjective2(chatId, bot) // Перезапуск блока 2
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
                        userPadezh.remove(chatId)
                        userStates.remove(chatId)
                        userColumnOrder.remove(chatId)
                        sendPadezhSelection(chatId, bot, tableFile)
                        println("Ж24 Сброс данных завершен: chatId = $chatId")
                    }
                    data == "next_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        // Пересчитываем состояние блоков перед проверкой
                        initializeUserBlockStates(chatId, tableFile)
                        val blockStates = userBlockCompleted[chatId] ?: Triple(false, false, false)
                        println("Ж25 Обработка 'Следующий блок': chatId = $chatId, currentBlock = $currentBlock")

                        userStates.remove(chatId)
                        userPadezh.remove(chatId)
                        userColumnOrder.remove(chatId)

                        when {
                            currentBlock == 1 && blockStates.first -> {
                                userBlocks[chatId] = 2
                                handleBlock2(chatId, bot, tableFile, wordUz, wordRus)
                            }
                            currentBlock == 2 && blockStates.second -> {
                                userBlocks[chatId] = 3
                                handleBlock3(chatId, bot, tableFile, wordUz, wordRus)
                            }
                            currentBlock == 3 && blockStates.third -> {
                                userBlocks[chatId] = 4
                                checkBlocksBeforeTest(chatId, bot, tableFile)
                            }
                            else -> {
                                val notCompletedBlocks = mutableListOf<String>()
                                if (!blockStates.first) notCompletedBlocks.add("Блок 1")
                                if (!blockStates.second) notCompletedBlocks.add("Блок 2")
                                if (!blockStates.third) notCompletedBlocks.add("Блок 3")

                                val messageText = "Вы не выполнили следующие блоки:\n" +
                                        notCompletedBlocks.joinToString("\n") +
                                        "\nПройдите их перед тестом."

                                bot.sendMessage(
                                    chatId = ChatId.fromId(chatId),
                                    text = messageText,
                                    replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                                        InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
                                    )
                                )
                                println("⚠️ Отправлено сообщение о незавершенных блоках: $notCompletedBlocks")
                            }
                        }
                    }

                    data == "prev_block" -> {
                        val currentBlock = userBlocks[chatId] ?: 1
                        println("Ж29 Обработка 'Предыдущий блок': chatId = $chatId, currentBlock = $currentBlock")
                        userStates.remove(chatId)
                        userPadezh.remove(chatId)

                        if (currentBlock > 1) {
                            userBlocks[chatId] = currentBlock - 1
                            handleBlock(chatId, bot, tableFile, wordUz, wordRus)
                            println("Ж30 Возврат на предыдущий блок: chatId = $chatId, prevBlock = ${userBlocks[chatId]}")
                        }
                    }
//                    data == "test_nouns" -> {
//                        println("Ж30 Запуск теста по существительным: chatId = $chatId")
//                        handleBlockTest(chatId, bot)
//                        println("Ж31 Вызвана функция handleBlockTest")
//                    }
//                    data == "adjective_block" -> {
//                        userStates[chatId] = 0
//                        println("Ж32 Переход к блоку прилагательных: chatId = $chatId")
//                        handleBlockAdjective1(chatId, bot)
//                        println("Ж33 Вызвана функция handleBlockAdjective")
//                    }
                }
            }
        }
    }

    bot.startPolling()
    println("Ж31 Бот начал опрос обновлений")
}
// Отправка приветственного сообщения
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
// Отправка стартового меню с выбором блока
fun sendStartMenu(chatId: Long, bot: Bot) {
    println("jjj sendStartMenu // Отправка стартового меню с выбором блока")

    val buttons = listOf(
        listOf(InlineKeyboardButton.CallbackData("Блок 1", "block:1")),
        listOf(InlineKeyboardButton.CallbackData("Блок 2", "block:2")),
        listOf(InlineKeyboardButton.CallbackData("Блок 3", "block:3")),
        listOf(InlineKeyboardButton.CallbackData("Тестовый блок", "block:test")),
        listOf(InlineKeyboardButton.CallbackData("Прилагательные 1", "block:adjective1")),
        listOf(InlineKeyboardButton.CallbackData("Прилагательные 2", "block:adjective2"))
    )

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Выберите блок для работы:",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
    println("jjj Стартовое меню отправлено пользователю: chatId=$chatId")
}



// Обработка основного блока для пользователя
fun handleBlock(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("AAA handleBlock // Обработка основного блока для пользователя")
    println("A1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")

    val currentBlock = userBlocks[chatId] ?: 1
    println("A2 Текущий блок пользователя: $currentBlock, userPadezh=${userPadezh[chatId]}")

    initializeUserBlockStates(chatId, filePath)
    println("A3 Состояния блоков пользователя после инициализации: ${userBlockCompleted[chatId]}")


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
        4 -> {
            println("A9 Запуск handleBlockTest для пользователя $chatId")
            handleBlockTest(chatId, bot)
        }
        5 -> {
            println("A10 Запуск handleBlockAdjective1 для пользователя $chatId")
            handleBlockAdjective1(chatId, bot)
        }
        6 -> {
            println("A11 Запуск handleBlockAdjective2 для пользователя $chatId")
            handleBlockAdjective2(chatId, bot)
        }
        else -> {
            println("A12 Неизвестный блок: $currentBlock для пользователя $chatId")
            bot.sendMessage(
                chatId = ChatId.fromId(chatId),
                text = "Неизвестный блок: $currentBlock"
            )
        }
    }
    println("A13 Выход из функции handleBlock")
}

// Обработка блока 1
fun handleBlock1(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("BBB handleBlock1 // Обработка блока 1")
    if (userPadezh[chatId] == null) {
        println("A4 Падеж не выбран для пользователя. Отправляем выбор падежа.")
        sendPadezhSelection(chatId, bot, filePath)
        return
    }
    println("A5 Падеж выбран: ${userPadezh[chatId]}")
    if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
        sendWordMessage(chatId, bot, tableFile)
        println("B0 Отправлена клавиатура для выбора слов: chatId = $chatId")
    } else {
        println("B1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")
        sendStateMessage(chatId, bot, tableFile, wordUz, wordRus)
        println("B2 Выход из функции handleBlock1")
    }
}

// Обработка блока 2
fun handleBlock2(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("CCC handleBlock2 // Обработка блока 2")

    // Если не выбран падеж — отправляем меню выбора
    if (userPadezh[chatId] == null) {
        sendPadezhSelection(chatId, bot, filePath)
        return
    }

    // Если не выбрано слово — отправляем меню выбора
    if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
        sendWordMessage(chatId, bot, filePath)
        return
    }

    // Определяем диапазоны падежей
    val blockRanges = mapOf(
        "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7"),
        "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14"),
        "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21"),
        "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28"),
        "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"),
        "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42", "F36-F42")
    )[userPadezh[chatId]] ?: return

    // Если порядок столбцов не сгенерирован — перемешиваем
    if (userColumnOrder[chatId].isNullOrEmpty()) {
        userColumnOrder[chatId] = blockRanges.shuffled().toMutableList()
    }

    val currentState = userStates[chatId] ?: 0
    val shuffledColumns = userColumnOrder[chatId]!!
    println("Текущее состояние: $currentState, Колонки: $shuffledColumns")

    // Если ВСЕ столбцы пройдены — отправляем финальное сообщение
    if (currentState >= shuffledColumns.size) {
        println("✅ Блок 2 завершен. Добавляем балл и отправляем финальное меню.")
        addScoreForPadezh(chatId, userPadezh[chatId].toString(), filePath, block = 2)
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
        return
    }

    val range = shuffledColumns[currentState]
    val messageText = generateMessageFromRange(filePath, "Существительные 2", range, wordUz, wordRus)
    val isLastMessage = currentState == shuffledColumns.size - 1

    // Отправляем сообщение
    if (isLastMessage) {
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = messageText, parseMode = ParseMode.MARKDOWN_V2)
        addScoreForPadezh(chatId, userPadezh[chatId].toString(), filePath, block = 2)
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

// Запуск блока 3 для пользователя
fun handleBlock3(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("🚀 handleBlock3 // Запуск блока 3 для пользователя $chatId")

    if (wordUz.isNullOrBlank() || wordRus.isNullOrBlank()) {
        sendWordMessage(chatId, bot, filePath)
        println("❌ Ошибка: слова не выбраны, запрошен повторный выбор.")
        return
    }

    val allRanges = listOf(
        "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7",
        "A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14",
        "A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21",
        "A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28",
        "A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"
    )

    // Если у пользователя еще нет порядка столбцов — перемешиваем
    if (userColumnOrder[chatId].isNullOrEmpty()) {
        userColumnOrder[chatId] = allRanges.shuffled().toMutableList()
        println("🔀 Перемешали мини-блоки для пользователя $chatId: ${userColumnOrder[chatId]}")
    }

    val shuffledRanges = userColumnOrder[chatId]!!
    val currentState = userStates[chatId] ?: 0
    println("📌 Текущее состояние пользователя: $currentState")

    // Загружаем баллы пользователя
    val completedRanges = getCompletedRanges(chatId, filePath)
    println("✅ Пройденные мини-блоки: $completedRanges")

    // Ищем первый неповторяющийся диапазон
    var currentRange: String? = null
    for (range in shuffledRanges) {
        if (!completedRanges.contains(range)) {
            currentRange = range
            break
        }
    }

    // Если ВСЕ 30 диапазонов пройдены, то берем текущий (изначально заданный) диапазон
    if (currentRange == null) {
        println("⚠️ ВСЕ 30 мини-блоков уже пройдены! Берем текущий из списка.")
        currentRange = shuffledRanges[currentState % shuffledRanges.size]
    }

    println("🎯 Выбран диапазон: $currentRange")

    val messageText = try {
        generateMessageFromRange(filePath, "Существительные 3", currentRange!!, wordUz, wordRus)
    } catch (e: Exception) {
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        println("❌ Ошибка при генерации сообщения: ${e.message}")
        return
    }

    val isLastRange = currentState == 5 // 6 сообщений всего
    println("📩 Отправка сообщения: \"$messageText\"")

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = messageText,
        parseMode = ParseMode.MARKDOWN_V2,
        replyMarkup = if (isLastRange) null else InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus")
        )
    )

    // Сохраняем прогресс
    saveUserProgressBlok3(chatId, filePath, currentRange!!)
    println("💾 Прогресс сохранен: $currentRange")

    if (!isLastRange) {
        println("📌 Переход к следующему шагу")
    } else {
        println("🏁 Мини-блоки завершены, вызываем финальное меню")
        updateUserProgressForMiniBlocks(chatId, filePath, shuffledRanges.indices.toList())
        sendFinalButtons(chatId, bot, wordUz, wordRus, filePath)
    }
}

// Обработка теста по существительным
fun handleBlockTest(chatId: Long, bot: Bot) {
    println("TTT handleBlockTest // Обработка теста по существительным")
    println("T1 Вход в функцию. Параметры: chatId=$chatId")

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Еще не реализовано.",
        replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Перейти к прилагательным", "block:adjective1")
        )
    )
    println("T2 Сообщение-заглушка отправлено пользователю $chatId")
}

// Переход к блоку с прилагательными
fun handleBlockAdjective1(chatId: Long, bot: Bot) {
    println("UUU handleBlockAdjective1 // Переход к блоку с прилагательными")
    println("U1 Вход в функцию. Параметры: chatId=$chatId")


    val currentBlock = userBlocks[chatId] ?: 1
    //val sheetName = if (currentBlock == 5) "Прилагательные 1" else "Прилагательные 2"
    println("U2 Текущий блок: $currentBlock, Лист: Прилагательные 1")


    // Инициализация замен, если их ещё нет
    if (userReplacements[chatId].isNullOrEmpty()) {
        initializeSheetColumnPairsFromFile(chatId)
        println("U3 Сгенерированы замены: ${userReplacements[chatId]} -------------------------------------------------------------------------")
    }


    val rangesForAdjectives = listOf(
        "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7",
        "F1-F7", "G1-G7", "H1-H7", "I1-I7", "J1-J7", "K1-K7", "L1-L7"
    )
    println("U4 Диапазоны мини-блоков: $rangesForAdjectives")

    println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  ${userStates[chatId]}")
    val currentState = userStates[chatId] ?: 0
    println("2!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  $currentState")
    println("U5 Текущее состояние пользователя: $currentState")

    // Если все мини-блоки завершены
    if (currentState >= rangesForAdjectives.size) {
        println("U6 Все мини-блоки завершены, отправляем финальное меню.")
        sendFinalButtonsForAdjectives(chatId, bot)
        userReplacements.remove(chatId) // Очищаем замены
        return
    }

    val currentRange = rangesForAdjectives[currentState]
    println("U7 Обрабатываем текущий диапазон: $currentRange")

    // Генерация сообщения с заменами
    val messageText = try {
        generateAdjectiveMessage(tableFile, "Прилагательные 1", currentRange, userReplacements[chatId]!!)
    } catch (e: Exception) {
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        println("U8 Ошибка при генерации сообщения: ${e.message}")
        return
    }
    println("U9 Сообщение сгенерировано: $messageText")

    val isLastRange = currentState == rangesForAdjectives.size - 1
    println("U10 Проверка на последний диапазон: $isLastRange")

    if (userStates[chatId] == null) { // Если пользователь только вошел в блок
        sendReplacementsMessage(chatId, bot) // Отправляем сообщение с 9 парами слов
    }


//    bot.sendMessage(
//        chatId = ChatId.fromId(chatId),
//        text = messageText,
//        parseMode = ParseMode.MARKDOWN_V2,
//        replyMarkup = if (isLastRange) null else InlineKeyboardMarkup.createSingleRowKeyboard(
//            InlineKeyboardButton.CallbackData("Далее", "next_adjective:$currentBlock")
//        )
//    )
//    println("U11 Сообщение отправлено пользователю: chatId=$chatId")

    // Отправляем сообщение (либо с кнопкой "Далее", либо финальное меню)
    if (isLastRange) {
        println("Последний диапазон. Отправляем финальное меню после этого сообщения.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2
        )
        sendFinalButtonsForAdjectives(chatId, bot)
    } else {
        println("Не последний диапазон. Отправляем сообщение с кнопкой 'Далее'.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2,
            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                InlineKeyboardButton.CallbackData("Далее", "next_adjective:$chatId")
            )
        )
    }

//    if (!isLastRange) {
//        userStates[chatId] = currentState + 1
//        println("U12 Обновлено состояние пользователя: chatId=$chatId, newState=${userStates[chatId]}")
//    }
}

// Переход к блоку с прилагательными
fun handleBlockAdjective2(chatId: Long, bot: Bot) {
    println("UUU2 handleBlockAdjective2 // Переход к блоку с прилагательными")
    println("U21 Вход в функцию. Параметры: chatId=$chatId")


    val currentBlock = userBlocks[chatId] ?: 1
    //val sheetName = if (currentBlock == 5) "Прилагательные 1" else "Прилагательные 2"
    println("U22 Текущий блок: $currentBlock, Лист: Прилагательные 2")


    // Инициализация замен, если их ещё нет
    if (userReplacements[chatId].isNullOrEmpty()) {
        initializeSheetColumnPairsFromFile(chatId)
        println("U23 Сгенерированы замены: ${userReplacements[chatId]}++++++++++++++++++++++++++++++++++++++++++++++++++++")
    }

    val rangesForAdjectives = listOf(
        "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7",
        "F1-F7", "G1-G7", "H1-H7", "I1-I7"
    )
    println("U24 Диапазоны мини-блоков: $rangesForAdjectives")

    val currentState = userStates[chatId] ?: 0
    println("U25 Текущее состояние пользователя: $currentState")

    // Если все мини-блоки завершены
    if (currentState >= rangesForAdjectives.size) {
        println("U26 Все мини-блоки завершены, отправляем финальное меню.")
        sendFinalButtonsForAdjectives(chatId, bot)
        userReplacements.remove(chatId) // Очищаем замены
        return
    }

    val currentRange = rangesForAdjectives[currentState]
    println("U27 Обрабатываем текущий диапазон: $currentRange")

    // Генерация сообщения с заменами
    val messageText = try {
        generateAdjectiveMessage(tableFile, "Прилагательные 2", currentRange, userReplacements[chatId]!!)
    } catch (e: Exception) {
        bot.sendMessage(chatId = ChatId.fromId(chatId), text = "Ошибка при формировании сообщения.")
        println("U28 Ошибка при генерации сообщения: ${e.message}")
        return
    }
    println("U29 Сообщение сгенерировано: $messageText")

    val isLastRange = currentState == rangesForAdjectives.size - 1
    println("U210 Проверка на последний диапазон: $isLastRange")

    if (userStates[chatId] == null) { // Если пользователь только вошел в блок
        sendReplacementsMessage(chatId, bot) // Отправляем сообщение с 9 парами слов
    }

    // Отправляем сообщение (либо с кнопкой "Далее", либо финальное меню)
    if (isLastRange) {
        println("Последний диапазон. Отправляем финальное меню после этого сообщения.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2
        )
        sendFinalButtonsForAdjectives(chatId, bot)
    } else {
        println("Не последний диапазон. Отправляем сообщение с кнопкой 'Далее'.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            parseMode = ParseMode.MARKDOWN_V2,
            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                InlineKeyboardButton.CallbackData("Далее", "next_adjective:$chatId")
            )
        )
    }
    println("U211 Сообщение отправлено пользователю: chatId=$chatId")

//    if (!isLastRange) {
//        userStates[chatId] = currentState + 1
//        println("U12 Обновлено состояние пользователя: chatId=$chatId, newState=${userStates[chatId]}")
//    }
}

// Генерация списка замен для пользователя
fun generateReplacements(chatId: Long) {
    println("VVV generateReplacements // Генерация списка замен для пользователя $chatId")

    // Проверяем, есть ли данные для пользователя
    val userPairs = sheetColumnPairs[chatId] ?: run {
        println("❌ Ошибка: Нет данных в sheetColumnPairs для пользователя $chatId")
        return
    }

    // Берем только ключи (узбекские слова)
    val keysList = userPairs.keys.toList()
    if (keysList.size < 9) {
        println("⚠️ Предупреждение: у пользователя $chatId недостаточно пар для замены.")
    }

    // Создаем изменяемый Map (MutableMap)
    val replacements = mutableMapOf<Int, String>()
    keysList.take(9).forEachIndexed { index, key ->
        replacements[index + 1] = key
    }

    // Присваиваем userReplacements для конкретного пользователя
    userReplacements[chatId] = replacements
    println("✅ Список замен обновлен для пользователя $chatId: $replacements")
}

// Генерация сообщения для блока прилагательных
fun generateAdjectiveMessage(
    filePath: String,
    sheetName: String,
    range: String,
    replacements: Map<Int, String>
): String {
    println("WWW generateAdjectiveMessage // Генерация сообщения для блока прилагательных")
    println("W1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName, range=$range")

    val rawText = generateMessageFromRange(filePath, sheetName, range, null, null)
    println("W2 Сырые данные из диапазона: \"$rawText\"")

    val processedText = rawText.replace(Regex("[1-9]")) { match ->
        val digit = match.value.toInt()
        replacements[digit] ?: match.value // Если нет замены, оставляем как есть
    }
    println("W3 Обработанный текст: \"$processedText\"")
    return processedText
}

// Отправка финального меню
fun sendFinalButtonsForAdjectives(chatId: Long, bot: Bot) {
    println("XXX sendFinalButtonsForAdjectives // Отправка финального меню")
    println("X1 Вход в функцию. Параметры: chatId=$chatId")

    val currentBlock = userBlocks[chatId] ?: 5  // По умолчанию 5 (прилагательные 1)
    val repeatCallback = if (currentBlock == 5) "block:adjective1" else "block:adjective2"
    val changeWordsCallback = if (currentBlock == 5) "change_words_adjective1" else "change_words_adjective2"

    val navigationButton = if (currentBlock == 5) {
        InlineKeyboardButton.CallbackData("Следующий блок", "block:adjective2")
    } else {
        InlineKeyboardButton.CallbackData("Предыдущий блок", "block:adjective1")
    }

    val buttons = listOf(
        listOf(InlineKeyboardButton.CallbackData("Повторить", repeatCallback)),
        listOf(InlineKeyboardButton.CallbackData("Изменить набор слов", changeWordsCallback)),
        listOf(InlineKeyboardButton.CallbackData("Начальное меню", "main_menu")),
        listOf(navigationButton)
    )

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим блоком прилагательных. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
    println("X2 Финальное меню отправлено пользователю: chatId=$chatId")
}

// Извлечение случайных слов из Excel
fun extractRandomWords(filePath: String, sheetName: String, column: Int, count: Int): List<String> {
    println("YYY extractRandomWords // Извлечение случайных слов из Excel")
    println("Y1 Вход в функцию. Параметры: filePath=$filePath, sheetName=$sheetName, column=$column, count=$count")

    val file = File(filePath)
    if (!file.exists()) {
        println("Y2 Ошибка: файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден")
    }

    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet(sheetName)
        ?: throw IllegalArgumentException("Лист $sheetName не найден")
    println("Y3 Лист найден: $sheetName")

    val rows = (1..sheet.lastRowNum).shuffled().take(count)
    val words = rows.mapNotNull { rowIndex ->
        val row = sheet.getRow(rowIndex)
        row?.getCell(column)?.toString()?.trim()
    }
    workbook.close()
    println("Y4 Извлеченные слова: $words")
    return words
}

// Формирование клавиатуры для выбора падежа
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

// Получение колонок текущего блока
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

// Чтение баллов пользователя для блока
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

// Формирование кнопок выбора падежей
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




// Генерация кнопки /start
fun generateUsersButton(): List<List<KeyboardButton>> {
    println("JJJ generateUsersButton // Генерация кнопки /start")
    println("J1 Вход в функцию.")
    val result = listOf(
        listOf(KeyboardButton("/start"))
    )
    println("J2 Сформированы кнопки: $result")
    return result
}


// Отправка клавиатуры с выбором слов
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


// Создание клавиатуры из Excel-файла
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


// Извлечение слов из callback data
fun extractWordsFromCallback(data: String): Pair<String, String> {
    println("MMM extractWordsFromCallback // Извлечение слов из callback data")
    println("M1 Вход в функцию. Параметры: data=$data")

    val content = data.substringAfter("word:(").substringBefore(")")
    val wordUz = content.substringBefore(" - ").trim()
    val wordRus = content.substringAfter(" - ").trim()

    println("M2 Извлеченные слова: wordUz=$wordUz, wordRus=$wordRus")
    return wordUz to wordRus
}


// Отправка сообщения по текущему состоянию и падежу
fun sendStateMessage(chatId: Long, bot: Bot, filePath: String, wordUz: String?, wordRus: String?) {
    println("NNN sendStateMessage // Отправка сообщения по текущему состоянию и падежу")
    println("N1 Вход в функцию. Параметры: chatId=$chatId, filePath=$filePath, wordUz=$wordUz, wordRus=$wordRus")

    val (selectedPadezh, rangesForPadezh, currentState) = validateUserState(chatId, bot) ?: return
    println("N2 Валидное состояние: selectedPadezh=$selectedPadezh, rangesForPadezh=$rangesForPadezh, currentState=$currentState")

    processStateAndSendMessage(chatId, bot, filePath, wordUz, wordRus, selectedPadezh, rangesForPadezh, currentState)
    println("N3 Завершение sendStateMessage")
}


// Проверка состояния пользователя
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

// Обработка состояния и отправка сообщения
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

// Отправка сообщения или переход к следующему шагу
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

// Отправка финального меню действий
fun sendFinalButtons(chatId: Long, bot: Bot, wordUz: String?, wordRus: String?, filePath: String) {
    println("RRR sendFinalButtons // Отправка финального меню действий")
    println("R1 Вход в функцию. Параметры: chatId=$chatId, wordUz=$wordUz, wordRus=$wordRus")

    val currentBlock = userBlocks[chatId] ?: 1
    println("R2 Текущий блок пользователя: $currentBlock")

    val buttons = mutableListOf(
        listOf(InlineKeyboardButton.CallbackData("Повторить", "repeat:$wordUz:$wordRus")),
        listOf(InlineKeyboardButton.CallbackData("Изменить слово", "change_word"))
    )

    // Если пользователь находится в третьем блоке, добавляем кнопку "Пройти тест по существительным"
    if (currentBlock == 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Пройти тест по существительным", "block:test")))
        println("R3 Добавлена кнопка 'Пройти тест по существительным'")
    } else {
        // Для других блоков оставляем кнопку "Изменить падеж"
        buttons.add(listOf(InlineKeyboardButton.CallbackData("Изменить падеж", "change_Padezh")))
        println("R4 Добавлена кнопка 'Изменить падеж'")
    }

    // Добавляем кнопки для переключения блоков
    if (currentBlock > 1) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
        println("R5 Добавлена кнопка '⬅️ Предыдущий блок'")
    }
    if (currentBlock < 3) {
        buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
        println("R6 Добавлена кнопка '➡️ Следующий блок'")
    }

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = "Вы завершили все этапы работы с этим словом и падежом. Что будем делать дальше?",
        replyMarkup = InlineKeyboardMarkup.create(buttons)
    )
    println("R7 Финальное меню отправлено пользователю $chatId")
}

// Генерация сообщения из диапазона Excel
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

// Экранирование Markdown V2
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

// Обработка узбекского слова в контексте строки
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

// Извлечение и обработка ячеек из диапазона
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

// Проверка состояния пользователя
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

// Сохранение данных в файл
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

// Добавляет балл для падежа
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

// Проверка состояния пользователя
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

// Начало добавления балла для пользователя
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

// Обработка содержимого ячейки
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

// Обработка ячейки без форматированных участков
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
    println("h1 Входные параметры: chatId=$chatId, filePath=$filePath")

    val block1Completed = checkUserState(chatId, filePath, block = 1)
    val block2Completed = checkUserState(chatId, filePath, block = 2)
    val block3Completed = checkUserStateBlock3(chatId, filePath) // 🔥 Исправили для блока 3

    println("h2 Состояние блока 1: $block1Completed")
    println("h3 Состояние блока 2: $block2Completed")
    println("h4 Состояние блока 3: $block3Completed (Пройдено мини-блоков: ${getCompletedRanges(chatId, filePath).size}/30)")

    userBlockCompleted[chatId] = Triple(block1Completed, block2Completed, block3Completed)
}

// Новая функция: Обновление прогресса пользователя для мини-блоков
fun updateUserProgressForMiniBlocks(chatId: Long, filePath: String, completedMiniBlocks: List<Int>) {
    println("iii updateUserProgressForMiniBlocks // Обновление прогресса пользователя по мини-блокам")

    println("i1 Входные параметры: chatId=$chatId, filePath=$filePath, completedMiniBlocks=$completedMiniBlocks")
    val file = File(filePath)
    if (!file.exists()) {
        println("i2 Ошибка: Файл $filePath не найден.")
        throw IllegalArgumentException("Файл $filePath не найден.")
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя")
        if (sheet == null) {
            println("i3 Ошибка: Лист 'Состояние пользователя' не найден.")
            throw IllegalArgumentException("Лист 'Состояние пользователя' не найден.")
        }

        // Ищем строку пользователя
        for (row in sheet) {
            val idCell = row.getCell(0)
            val chatIdFromCell = when (idCell?.cellType) {
                CellType.NUMERIC -> idCell.numericCellValue.toLong()
                CellType.STRING -> idCell.stringCellValue.toDoubleOrNull()?.toLong()
                else -> null
            }

            if (chatIdFromCell == chatId) {
                println("i4 Пользователь найден. Обновляем прогресс.")

                // Обновляем ячейки от L вправо
                completedMiniBlocks.forEach { miniBlock ->
                    val columnIndex = 11 + miniBlock // L = 11, M = 12 и т.д.
                    val cell = row.getCell(columnIndex) ?: row.createCell(columnIndex)
                    val currentValue = cell.numericCellValue.takeIf { it > 0 } ?: 0.0
                    cell.setCellValue(currentValue + 1)
                    println("i5 Мини-блок $miniBlock: старое значение = $currentValue, новое значение = ${currentValue + 1}")
                }

                safelySaveWorkbook(workbook, filePath)
                println("i6 Прогресс успешно обновлен для пользователя $chatId.")
                return
            }
        }

        println("i7 Ошибка: Пользователь $chatId не найден в таблице.")
    }
}

// Отправка сообщения с 9 парами слов из sheetColumnPairs
fun sendReplacementsMessage(chatId: Long, bot: Bot) {
    println("### sendReplacementsMessage // Отправка сообщения с 9 парами слов из sheetColumnPairs")

    val userPairs = sheetColumnPairs[chatId]

    if (userPairs.isNullOrEmpty()) {
        println("❌ Ошибка: Данные в sheetColumnPairs отсутствуют для пользователя $chatId.")
        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = "Ошибка: Данные для отправки не найдены."
        )
        return
    }

    // Формируем сообщение в нужном формате
    val messageText = userPairs.entries.joinToString("\n") { (key, value) ->
        "$key - $value"
    }

    bot.sendMessage(
        chatId = ChatId.fromId(chatId),
        text = messageText
    )

    println("✅ Сообщение отправлено пользователю $chatId:\n$messageText")
}

// Сохранение прогресса пользователя
fun saveUserProgressBlok3(chatId: Long, filePath: String, range: String) {
    println("📌 saveUserProgressBlok3 // Сохранение прогресса пользователя")

    val file = File(filePath)
    if (!file.exists()) {
        println("❌ Ошибка: Файл $filePath не найден.")
        return
    }

    // Карта соответствий диапазонов столбцам
    val columnMapping = mapOf(
        "A1-A7" to 1, "B1-B7" to 2, "C1-C7" to 3, "D1-D7" to 4, "E1-E7" to 5, "F1-F7" to 6,
        "A8-A14" to 7, "B8-B14" to 8, "C8-C14" to 9, "D8-D14" to 10, "E8-E14" to 11, "F8-F14" to 12,
        "A15-A21" to 13, "B15-B21" to 14, "C15-C21" to 15, "D15-D21" to 16, "E15-E21" to 17, "F15-F21" to 18,
        "A22-A28" to 19, "B22-B28" to 20, "C22-C28" to 21, "D22-D28" to 22, "E22-E28" to 23, "F22-F28" to 24,
        "A29-A35" to 25, "B29-B35" to 26, "C29-C35" to 27, "D29-D35" to 28, "E29-E35" to 29, "F29-F35" to 30
    )

    // Определяем, в какую колонку писать
    val columnIndex = columnMapping[range]
    if (columnIndex == null) {
        println("❌ Ошибка: Диапазон $range не найден в карте соответствий.")
        return
    }
    println("✅ Диапазон $range -> Записываем в колонку $columnIndex")

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя 3 блок") ?: workbook.createSheet("Состояние пользователя 3 блок")

        // Поиск пользователя в первой строке (идём вправо)
        val headerRow = sheet.getRow(0) ?: sheet.createRow(0)
        var userColumnIndex = -1

        for (col in 0 until 31) {  // Проверяем до 30 столбца
            val cell = headerRow.getCell(col) ?: headerRow.createCell(col)
            if (cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                userColumnIndex = col
                break
            } else if (cell.cellType == CellType.BLANK) {  // Если нашли пустой, записываем ID
                cell.setCellValue(chatId.toDouble())
                userColumnIndex = col
                break
            }
        }

        if (userColumnIndex == -1) {
            println("⚠️ Нет свободного места для записи ID!")
            return
        }

        // Получаем строку для записи балла (по индексу столбца)
        val row = sheet.getRow(columnIndex) ?: sheet.createRow(columnIndex)
        val cell = row.getCell(userColumnIndex) ?: row.createCell(userColumnIndex)

        // Увеличиваем балл на 1
        val currentScore = cell.numericCellValue.takeIf { it > 0 } ?: 0.0
        cell.setCellValue(currentScore + 1)

        println("✅ Прогресс обновлён: chatId=$chatId, столбец=$userColumnIndex, строка=$columnIndex, новый балл=${currentScore + 1}")

        safelySaveWorkbook(workbook, filePath)
    }
}

// Проверка пройденных мини-блоков
fun getCompletedRanges(chatId: Long, filePath: String): Set<String> {
    println("📊 getCompletedRanges // Проверка пройденных мини-блоков")

    val completedRanges = mutableSetOf<String>()
    val file = File(filePath)
    if (!file.exists()) {
        println("❌ Ошибка: Файл $filePath не найден.")
        return emptySet()
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя 3 блок") ?: return emptySet()
        val headerRow = sheet.getRow(0) ?: return emptySet()

        var userColumnIndex: Int? = null
        for (col in 0 until 31) {
            val cell = headerRow.getCell(col)
            if (cell != null && cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                userColumnIndex = col
                break
            }
        }

        if (userColumnIndex == null) {
            println("⚠️ Пользователь $chatId не найден в таблице прогресса.")
            return emptySet()
        }

        val rangeMapping = listOf(
            "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7",
            "A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14",
            "A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21",
            "A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28",
            "A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"
        )

        for ((index, range) in rangeMapping.withIndex()) {
            val row = sheet.getRow(index + 1) ?: continue
            val cell = row.getCell(userColumnIndex)
            val value = cell?.numericCellValue ?: 0.0
            if (value > 0) {
                completedRanges.add(range)
            }
        }
    }

    println("✅ Завершенные диапазоны для пользователя $chatId: $completedRanges")
    return completedRanges
}

// Проверка выполнения блоков перед тестом
fun checkBlocksBeforeTest(chatId: Long, bot: Bot, filePath: String) {
    println("🚦 checkBlocksBeforeTest // Проверка выполнения блоков перед тестом")

    // Загружаем состояние пользователя по блокам
    initializeUserBlockStates(chatId, filePath)
    val (block1Completed, block2Completed, block3Completed) = userBlockCompleted[chatId] ?: Triple(false, false, false)

    println("📊 Состояние блоков для пользователя $chatId:")
    println("✅ Блок 1: $block1Completed")
    println("✅ Блок 2: $block2Completed")
    println("✅ Блок 3: $block3Completed")

    // Проверяем, пройдены ли все блоки
    if (block1Completed && block2Completed && block3Completed) {
        println("✅ Все блоки завершены. Запускаем тест.")
        handleBlockTest(chatId, bot)
    } else {
        val notCompletedBlocks = mutableListOf<String>()
        if (!block1Completed) notCompletedBlocks.add("Блок 1")
        if (!block2Completed) notCompletedBlocks.add("Блок 2")
        if (!block3Completed) notCompletedBlocks.add("Блок 3")

        val messageText = "Вы не выполнили следующие блоки:\n" + notCompletedBlocks.joinToString("\n") +
                "\nПройдите их перед тестом."

        bot.sendMessage(
            chatId = ChatId.fromId(chatId),
            text = messageText,
            replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
            )
        )
        println("⚠️ Отправлено сообщение о незавершенных блоках: $notCompletedBlocks")
    }
}

//Проверка прогресса пользователя в блоке 3
fun checkUserStateBlock3(chatId: Long, filePath: String): Boolean {
    println("🛠 Проверяем прогресс пользователя в блоке 3")

    val file = File(filePath)
    if (!file.exists()) {
        println("❌ Файл не найден: $filePath")
        return false
    }

    WorkbookFactory.create(file).use { workbook ->
        val sheet = workbook.getSheet("Состояние пользователя 3 блок")
        if (sheet == null) {
            println("❌ Лист 'Состояние пользователя 3 блок' отсутствует.")
            return false
        }

        val headerRow = sheet.getRow(0) ?: return false

        // Находим колонку пользователя
        var userColumnIndex: Int? = null
        for (col in 0..30) {  // Проверяем ячейки от 2 по 31
            val cell = headerRow.getCell(col)
            if (cell != null && cell.cellType == CellType.NUMERIC && cell.numericCellValue.toLong() == chatId) {
                userColumnIndex = col
                break
            }
        }

        if (userColumnIndex == null) {
            println("⚠️ Пользователь $chatId не найден в таблице прогресса.")
            return false
        }

        // Проверяем, есть ли баллы во ВСЕХ 30 мини-блоках
        for (rowIdx in 1..30) { // Проверяем строки 2-31 (т.к. индексация с 0)
            val row = sheet.getRow(rowIdx) ?: return false
            val cell = row.getCell(userColumnIndex)
            val value = cell?.numericCellValue ?: 0.0
            if (value == 0.0) {
                println("❌ Мини-блок $rowIdx не завершен")
                return false
            }
        }

        println("✅ Все 30 мини-блоков завершены!")
        return true
    }
}

// Инициализация пар для пользователя
fun initializeSheetColumnPairsFromFile(chatId: Long) {
    println("### initializeSheetColumnPairsFromFile // Инициализация пар для пользователя $chatId")

    sheetColumnPairs[chatId] = mutableMapOf()  // Создаем пустой мап для пользователя
    val file = File(tableFile)
    if (!file.exists()) {
        println("❌ Файл $tableFile не найден!")
        return
    }

    val workbook = WorkbookFactory.create(file)
    val sheetNames = listOf("Существительные", "Глаголы", "Прилагательные")
    val userPairs = mutableMapOf<String, String>()  // Временный мап для хранения пар

    for (sheetName in sheetNames) {
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            println("⚠️ Лист $sheetName не найден!")
            continue  // ✅ Просто continue, без run {}
        }

        val candidates = mutableListOf<Pair<String, String>>() // Кандидаты на выборку

        for (i in 0..sheet.lastRowNum) {
            val row = sheet.getRow(i) ?: continue
            val key = row.getCell(0)?.toString()?.trim() ?: ""
            val value = row.getCell(1)?.toString()?.trim() ?: ""
            if (key.isNotEmpty() && value.isNotEmpty()) {
                candidates.add(key to value)
            }
        }

        if (candidates.size < 3) {
            println("⚠️ Недостаточно данных в листе $sheetName для выбора 3 пар.")
            continue
        }

        val selectedPairs = candidates.shuffled().take(3)  // Берем 3 случайные пары
        for ((key, value) in selectedPairs) {
            userPairs[key] = value
        }
    }

    workbook.close()

    if (userPairs.size == 9) {
        sheetColumnPairs[chatId] = userPairs  // Сохраняем пары в общий мап
        println("✅ Успешно загружены 9 пар для пользователя $chatId: $userPairs")
    } else {
        println("❌ Ошибка: Получено ${userPairs.size} пар вместо 9.")
    }
    generateReplacements(chatId)
}
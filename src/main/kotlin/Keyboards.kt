package com.github.kotlintelegrambot.keyboards

import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.KeyboardReplyMarkup
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.entities.keyboard.KeyboardButton


object Keyboards {

    // Клавиатура для команды /start
    fun getStartButton(): KeyboardReplyMarkup {
        val buttons = listOf(
            listOf(KeyboardButton("/start"))
        )
        return KeyboardReplyMarkup(keyboard = buttons, resizeKeyboard = true)
    }

    // Стартовое меню для выбора блока
    fun startMenu(): InlineKeyboardMarkup {
        val buttons = listOf(
            listOf(InlineKeyboardButton.CallbackData("Существительные 1", "nouns1")),
            listOf(InlineKeyboardButton.CallbackData("Существительные 2", "nouns2")),
            listOf(InlineKeyboardButton.CallbackData("Существительные 3", "nouns3")),
            listOf(InlineKeyboardButton.CallbackData("Тест", "block:test")),
            listOf(InlineKeyboardButton.CallbackData("Прилагательные 1", "adjective1")),
            listOf(InlineKeyboardButton.CallbackData("Прилагательные 2", "adjective2")),
            listOf(InlineKeyboardButton.CallbackData("Глаголы 1", "verbs1")),
            listOf(InlineKeyboardButton.CallbackData("Глаголы 2", "verbs2")),
            listOf(InlineKeyboardButton.CallbackData("Глаголы 3", "verbs3"))
        )
        return InlineKeyboardMarkup.create(buttons)
    }

    // Клавиатура для выбора падежа
    fun padezhSelection(currentBlock: Int, padezhColumns: Map<String, Int>, userScores: Map<String, Int>): InlineKeyboardMarkup {
        val buttons = padezhColumns.keys.map { padezh ->
            val score = userScores[padezh] ?: 0
            InlineKeyboardButton.CallbackData("$padezh [$score]", "Padezh:$padezh")
        }.map { listOf(it) }.toMutableList()

        if (currentBlock > 1) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
        }
        if (currentBlock < 3) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
        }
        return InlineKeyboardMarkup.create(buttons)
    }

    // Финальное меню для блоков прилагательных
    fun finalAdjectiveButtons(currentBlock: Int): InlineKeyboardMarkup {
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
            //listOf(InlineKeyboardButton.CallbackData("Начальное меню", "main_menu")),
            listOf(navigationButton)
        )
        return InlineKeyboardMarkup.create(buttons)
    }

    // Финальное меню для остальных блоков
    fun finalButtons(wordUz: String?, wordRus: String?, currentBlock: Int): InlineKeyboardMarkup {
        val buttons = mutableListOf(
            listOf(InlineKeyboardButton.CallbackData("Повторить", "repeat:$wordUz:$wordRus")),
            listOf(InlineKeyboardButton.CallbackData("Изменить слово", "change_word"))
        )
        if (currentBlock == 3) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("Пройти тест по существительным", "block:test")))
        } else {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("Изменить падеж", "change_Padezh")))
        }
        if (currentBlock > 1) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("⬅️ Предыдущий блок", "prev_block")))
        }
        if (currentBlock < 3) {
            buttons.add(listOf(InlineKeyboardButton.CallbackData("➡️ Следующий блок", "next_block")))
        }
        return InlineKeyboardMarkup.create(buttons)
    }

    // Клавиатура с одной кнопкой "Далее"
    fun nextButton(wordUz: String?, wordRus: String?): InlineKeyboardMarkup {
        val buttons = listOf(
            listOf(InlineKeyboardButton.CallbackData("Выбранное слово: $wordRus - $wordUz", "nouns1Change_word_random")),
            //listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")),
            listOf(InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus"))
        )
        return InlineKeyboardMarkup.create(buttons)
    }



    // Клавиатура для перехода к блоку прилагательных (например, из теста)
    fun goToAdjectivesButton(): InlineKeyboardMarkup {
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Перейти к прилагательным", "block:adjective1")
        )
    }

    // Клавиатура с кнопкой "Вернуться к блокам"
    fun returnToBlocksButton(): InlineKeyboardMarkup {
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
        )
    }

    // Пример клавиатуры «Далее» специально для глаголов
    fun nextVerbsButton(): InlineKeyboardMarkup {
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Далее", "next_verbs1")
        )
    }

    // Кнопка-заглушка в конце блока
    fun finalVerbsButton(): InlineKeyboardMarkup {
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("В главное меню", "main_menu")
        )
    }

    fun nextButtonWithHintToggle(wordUz: String?, wordRus: String?, isHintVisible: Boolean, blockId: String): InlineKeyboardMarkup {
        val baseButtons = listOf(
            listOf(InlineKeyboardButton.CallbackData("Выбранное слово: $wordRus - $wordUz", "nouns1Change_word_random")),
            //listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")),
            listOf(InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus"))
        )
        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        // Формат callback: toggleHint:{текущее_состояние}:{blockId}:{wordUz}:{wordRus}
        val toggleCallback = "toggleHint:${isHintVisible}:$blockId:$wordUz:$wordRus"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))
        return InlineKeyboardMarkup.create(baseButtons + listOf(toggleButtonRow))
    }

    fun nextPadezhButton(nextPadezh: String): InlineKeyboardMarkup {
        val button = InlineKeyboardButton.CallbackData("Выбрать падеж", "nouns1")
        return InlineKeyboardMarkup.createSingleRowKeyboard(button)
    }

    fun nextCaseButtonWithHintToggle(
        wordUz: String?,
        wordRus: String?,
        isHintVisible: Boolean,
        blockId: String,
        nextPadezh: String
    ): InlineKeyboardMarkup {
        val selectedWordRow = listOf(
            InlineKeyboardButton.CallbackData("Выбранное слово: $wordRus - $wordUz", "nouns1Change_word_random")
        )
        val nextCaseRow = listOf(
            InlineKeyboardButton.CallbackData("Выбрать падеж", "nouns1")
        )
        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        val toggleCallback = "toggleHint:${isHintVisible}:$blockId:$wordUz:$wordRus"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))
        return InlineKeyboardMarkup.create(listOf(selectedWordRow, nextCaseRow, toggleButtonRow))
    }

//    fun transitionToNouns2ButtonWithHintToggle(
//        wordUz: String?,
//        wordRus: String?,
//        isHintVisible: Boolean,
//        blockId: String
//    ): InlineKeyboardMarkup {
//        val selectedWordRow = listOf(
//            InlineKeyboardButton.CallbackData("Выбранное слово: $wordRus - $wordUz", "nouns1Change_word_random")
//        )
//        val transitionRow = listOf(
//            InlineKeyboardButton.CallbackData("Перейти к Существительные 2", "nouns2")
//        )
//        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
//        val toggleCallback = "toggleHint:${isHintVisible}:$blockId:$wordUz:$wordRus"
//        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))
//        return InlineKeyboardMarkup.create(listOf(selectedWordRow, transitionRow, toggleButtonRow))
//    }
//
//    fun caseSelectionButton(): InlineKeyboardMarkup {
//        return InlineKeyboardMarkup.createSingleRowKeyboard(
//            InlineKeyboardButton.CallbackData("Выбор подежа", "nouns1")
//        )
//    }
}

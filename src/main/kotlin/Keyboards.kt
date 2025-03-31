package com.github.kotlintelegrambot.keyboards

import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.KeyboardReplyMarkup
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.entities.keyboard.KeyboardButton


object Keyboards {

    // Стартовое меню для выбора блока
    fun startMenu(): InlineKeyboardMarkup {
        println("Keyboards startMenu")
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

    // Финальное меню для остальных блоков
    fun finalButtons(wordUz: String?, wordRus: String?, currentBlock: Int): InlineKeyboardMarkup {
        println("Keyboards finalButtons")
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
        println("Keyboards nextButton")
        val buttons = listOf(
            listOf(InlineKeyboardButton.CallbackData("Сменить слово", "nouns1Change_word_random")),
            //listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")),
            listOf(InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus"))
        )
        return InlineKeyboardMarkup.create(buttons)
    }

    // Клавиатура для перехода к блоку прилагательных (например, из теста)
    fun goToAdjectivesButton(): InlineKeyboardMarkup {
        println("Keyboards goToAdjectivesButton")
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Перейти к прилагательным", "block:adjective1")
        )
    }

    // Клавиатура с кнопкой "Вернуться к блокам"
    fun returnToBlocksButton(): InlineKeyboardMarkup {
        println("Keyboards returnToBlocksButton")
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
        )
    }

    // Пример клавиатуры «Далее» специально для глаголов
    fun nextVerbsButton(): InlineKeyboardMarkup {
        println("Keyboards nextVerbsButton")
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("Далее", "next_verbs1")
        )
    }

    // Кнопка-заглушка в конце блока
    fun finalVerbsButton(): InlineKeyboardMarkup {
        println("Keyboards finalVerbsButton")
        return InlineKeyboardMarkup.createSingleRowKeyboard(
            InlineKeyboardButton.CallbackData("В главное меню", "main_menu")
        )
    }

    fun nextButtonWithHintToggle(wordUz: String?, wordRus: String?, isHintVisible: Boolean, blockId: String): InlineKeyboardMarkup {
        println("Keyboards nextButtonWithHintToggle")
        val baseButtons = listOf(
            listOf(InlineKeyboardButton.CallbackData("Сменить слово", "nouns1Change_word_random")),
            //listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")),
            listOf(InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus"))
        )
        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        // Формат callback: toggleHint:{текущее_состояние}:{blockId}:{wordUz}:{wordRus}
        val toggleCallback = "toggleHint:${isHintVisible}:$blockId:$wordUz:$wordRus"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))
        return InlineKeyboardMarkup.create(baseButtons + listOf(toggleButtonRow))
    }

    fun nextCaseButtonWithHintToggleNouns1(
        wordUz: String?,
        wordRus: String?,
        isHintVisible: Boolean,
        blockId: String,
    ): InlineKeyboardMarkup {
        println("Keyboards nextCaseButtonWithHintToggle")
        val selectedWordRow = listOf(
            InlineKeyboardButton.CallbackData("Сменить слово", "nouns1Change_word_random")
        )
        val nextCaseRow = listOf(
            InlineKeyboardButton.CallbackData("Выбрать падеж", "nouns1")
        )
        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        val toggleCallback = "toggleHint:${isHintVisible}:$blockId:$wordUz:$wordRus"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))
        return InlineKeyboardMarkup.create(listOf(selectedWordRow, nextCaseRow, toggleButtonRow))
    }

    fun nextCaseButtonWithHintToggleNouns2(
        wordUz: String?,
        wordRus: String?,
    ): InlineKeyboardMarkup {
        println("Keyboards nextCaseButtonWithHintToggle")
        val selectedWordRow = listOf(
            InlineKeyboardButton.CallbackData("Сменить слово", "nouns1Change_word_random")
        )
        val nextCaseRow = listOf(
            InlineKeyboardButton.CallbackData("Выбрать падеж", "nouns2")
        )
        return InlineKeyboardMarkup.create(listOf(selectedWordRow, nextCaseRow))
    }

    fun buttonNextChengeWord(wordUz: String?, wordRus: String?): InlineKeyboardMarkup {
        println("Keyboards nextButton")
        val buttons = listOf(
            listOf(InlineKeyboardButton.CallbackData("Сменить слово", "nouns3Change_word_random")),
            //listOf(InlineKeyboardButton.CallbackData("Меню", "main_menu")),
            listOf(InlineKeyboardButton.CallbackData("Далее", "next:$wordUz:$wordRus"))
        )
        return InlineKeyboardMarkup.create(buttons)
    }

    fun adjective1HintToggleKeyboard(isHintVisible: Boolean, isLast: Boolean): InlineKeyboardMarkup {
        println("Keyboards adjective1HintToggleKeyboard")
        val baseButtons = mutableListOf(
            listOf(InlineKeyboardButton.CallbackData("Сменить набор слов", "change_words_adjective1"))
        )

        val navigationButton = if (isLast) {
            InlineKeyboardButton.CallbackData("Следующий блок", "adjective2")
        } else {
            InlineKeyboardButton.CallbackData("Далее", "next:заглушкаbola:заглушкаребенок")
        }

        baseButtons.add(listOf(navigationButton))

        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        val toggleCallback = "toggleHintAdjective1:$isHintVisible"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))

        return InlineKeyboardMarkup.create(baseButtons + listOf(toggleButtonRow))
    }

    fun adjective2HintToggleKeyboard(isHintVisible: Boolean, isLast: Boolean): InlineKeyboardMarkup {
        println("Keyboards adjective2HintToggleKeyboard")
        val baseButtons = mutableListOf(
            listOf(InlineKeyboardButton.CallbackData("Сменить набор слов", "change_words_adjective2"))
        )

        val navigationButton = if (isLast) {
            InlineKeyboardButton.CallbackData("Перейти к глаголам", "verbs1")
        } else {
            InlineKeyboardButton.CallbackData("Далее", "next:заглушкаbola:заглушкаребенок")
        }

        baseButtons.add(listOf(navigationButton))

        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        val toggleCallback = "toggleHintAdjective2:$isHintVisible"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))

        return InlineKeyboardMarkup.create(baseButtons + listOf(toggleButtonRow))
    }

    fun verbs1HintToggleKeyboard(isHintVisible: Boolean, isLast: Boolean): InlineKeyboardMarkup {
        println("Keyboards verbs1HintToggleKeyboard")
        val baseButtons = mutableListOf(
            listOf(InlineKeyboardButton.CallbackData("Сменить слово", "change_words_verbs1"))
        )

        val navigationButton = if (isLast) {
            InlineKeyboardButton.CallbackData("Перейти к следующему блоку", "verbs2")
        } else {
            InlineKeyboardButton.CallbackData("Далее", "next_verbs1")
        }

        baseButtons.add(listOf(navigationButton))

        val toggleText = if (isHintVisible) "Скрыть подсказку" else "Показать подсказку"
        val toggleCallback = "toggleHintVerbs1:$isHintVisible"
        val toggleButtonRow = listOf(InlineKeyboardButton.CallbackData(toggleText, toggleCallback))

        return InlineKeyboardMarkup.create(baseButtons + listOf(toggleButtonRow))
    }
}

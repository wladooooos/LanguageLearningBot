package com.github.kotlintelegrambot.config

import com.github.kotlintelegrambot.dispatcher.userWords

object Config {
    // Токен для подключения к Telegram API
    const val BOT_TOKEN: String = "7856005284:AAFVvPnRadWhaotjUZOmFyDFgUHhZ0iGsCo"

    // Путь к файлу таблицы Excel
    const val TABLE_FILE: String = "Table.xlsx"

    // Диапазоны для падежей (используются для формирования сообщений)
    val PADEZH_RANGES: Map<String, List<String>> = mapOf(
        "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7"),
        "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14"),
        "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21"),
        "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28"),
        "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35"),
        "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42")
    )

    // Диапазоны колонок для каждого блока (используются для формирования клавиатуры выбора падежей)
    val COLUMN_RANGES: Map<Int, Map<String, Int>> = mapOf(
        1 to mapOf(
            "Именительный" to 1,
            "Родительный" to 2,
            "Винительный" to 3,
            "Дательный" to 4,
            "Местный" to 5,
            "Исходный" to 6
        ),
        2 to mapOf(
            "Именительный" to 7,
            "Родительный" to 8,
            "Винительный" to 9,
            "Дательный" to 10,
            "Местный" to 11,
            "Исходный" to 12
        ),
        3 to mapOf(
            "Именительный" to 13,
            "Родительный" to 14,
            "Винительный" to 15,
            "Дательный" to 16,
            "Местный" to 17,
            "Исходный" to 18
        )
    )
}

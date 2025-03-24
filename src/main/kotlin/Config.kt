package com.github.kotlintelegrambot.config

object Config {
    // Токен для подключения к Telegram API
    //const val BOT_TOKEN: String = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE"
    const val BOT_TOKEN: String = "7856005284:AAFVvPnRadWhaotjUZOmFyDFgUHhZ0iGsCo"

    // Путь к файлу таблицы Excel
    //const val TABLE_FILE: String = "/home/ec2-user/Table.xlsx"
    const val TABLE_FILE: String = "Table.xlsx"

    //const val GVERBS_FILE = "/home/ec2-user/Глаголы.xlsx"
    const val GVERBS_FILE = "Глаголы.xlsx"

    // Диапазоны для падежей (используются для формирования сообщений)
    val PADEZH_RANGES: Map<String, List<String>> = mapOf(
        "Именительный" to listOf("A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7"),
        "Родительный" to listOf("A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14"),
        "Винительный" to listOf("A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21"),
        "Дательный" to listOf("A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28"),
        "Местный" to listOf("A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35"),
        "Исходный" to listOf("A36-A42", "B36-B42", "C36-C42", "D36-D42", "E36-E42")
    )

    // Все диапазоны для блока 3 (мини-блоки)
    val ALL_RANGES_BLOCK_3: List<String> = listOf(
        "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7", "F1-F7",
        "A8-A14", "B8-B14", "C8-C14", "D8-D14", "E8-E14", "F8-F14",
        "A15-A21", "B15-B21", "C15-C21", "D15-D21", "E15-E21", "F15-F21",
        "A22-A28", "B22-B28", "C22-C28", "D22-D28", "E22-E28", "F22-F28",
        "A29-A35", "B29-B35", "C29-C35", "D29-D35", "E29-E35", "F29-F35"
    )

    // Диапазоны для блока прилагательных 1
    val ADJECTIVE_RANGES_1: List<String> = listOf(
        "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7",
        "F1-F7", "G1-G7", "H1-H7", "I1-I7", "J1-J7", "K1-K7", "L1-L7"
    )

    // Диапазоны для блока прилагательных 2
    val ADJECTIVE_RANGES_2: List<String> = listOf(
        "A1-A7", "B1-B7", "C1-C7", "D1-D7", "E1-E7",
        "F1-F7", "G1-G7", "H1-H7", "I1-I7"
    )

    // Диапазоны колонок для проверки состояния пользователя по блокам
    val USER_STATE_BLOCK_RANGES: Map<Int, IntRange> = mapOf(
        1 to (1..6),
        2 to (7..12),
        3 to (13..18)
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

    // Карта соответствий диапазонов столбцам
    val COLUMN_MAPPING = mapOf(
        "A1-A7" to 1, "B1-B7" to 2, "C1-C7" to 3, "D1-D7" to 4, "E1-E7" to 5, "F1-F7" to 6,
        "A8-A14" to 7, "B8-B14" to 8, "C8-C14" to 9, "D8-D14" to 10, "E8-E14" to 11, "F8-F14" to 12,
        "A15-A21" to 13, "B15-B21" to 14, "C15-C21" to 15, "D15-D21" to 16, "E15-E21" to 17, "F15-F21" to 18,
        "A22-A28" to 19, "B22-B28" to 20, "C22-C28" to 21, "D22-D28" to 22, "E22-E28" to 23, "F22-F28" to 24,
        "A29-A35" to 25, "B29-B35" to 26, "C29-C35" to 27, "D29-D35" to 28, "E29-E35" to 29, "F29-F35" to 30
    )

    val GVERBS_RANGES_3 = listOf(
        "A1-A5", "B1-B5", "C1-C5", "D1-D5", "E1-E5", "F1-F5",
        "G1-G4",
        "H1-H5", "I1-I5", "J1-J5", "K1-K5", "L1-L5",
        "M1-M7", "N1-N7", "O1-O7", "P1-P7", "Q1-Q7", "R1-R7",
        "S1-S4", "T1-T4", "U1-U4", "V1-V4", "W1-W4", "X1-X4",
        "Y1-Y5", "Z1-Z5", "AA1-AA5", "AB1-AB5", "AC1-AC5", "AD1-AD5",
        "AE1-AE4", "AF1-AF4", "AG1-AG4", "AH1-AH4", "AI1-AI4", "AJ1-AJ4",

        "A9-A13", "B9-B13", "C9-C13", "D9-D13", "E9-E13", "F9-F13",
        "G9-G12", "H9-H12", "I9-I12", "J9-J12", "K9-K12", "L9-L12",
        "M9-M13", "N9-N13", "O9-O13", "P9-P13", "Q9-Q13", "R9-R13",
        "S9-S12", "T9-T12", "U9-U12", "V9-V12", "W9-W12", "X9-X12",
        "Y9-Y13", "Z9-Z13", "AA9-AA13", "AB9-AB13", "AC9-AC13", "AD9-AD13",
        "AE9-AE12", "AF9-AF12", "AG9-AG12", "AH9-AH12", "AI9-AI12", "AJ9-AJ12"
    )
}

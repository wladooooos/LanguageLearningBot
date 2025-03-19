object Globals {
    val userStates = mutableMapOf<Long, Int>() //Навигация внутри падежа
    val userPadezh = mutableMapOf<Long, String>() // Хранение выбранного падежа для каждого пользователя
    val userWords = mutableMapOf<Long, Pair<String, String>>() // Хранение выбранного слова для каждого пользователя
    val userBlocks = mutableMapOf<Long, Int>() // Хранение текущего блока для каждого пользователя
    val userBlockCompleted = mutableMapOf<Long, Triple<Boolean, Boolean, Boolean>>() // Состояния блоков (пройдено или нет)
    val userColumnOrder = mutableMapOf<Long, MutableList<String>>() // Для хранения случайного порядка столбцов
    val userWordUz: MutableMap<Long, String> = mutableMapOf()
    var userWordRus: MutableMap<Long, String> = mutableMapOf()
    val userReplacements = mutableMapOf<Long, Map<Int, String>>() // Хранение замен для чисел (1-9) для каждого пользователя
    var sheetColumnPairs = mutableMapOf<Long, Map<String, String>>() // Глобальная переменная для хранения пар лист/столбец (ключ и значение — строки)
    val userRandomVerb = mutableMapOf<Long, String>()
    val userMainMessageId = mutableMapOf<Long, Int>()
}
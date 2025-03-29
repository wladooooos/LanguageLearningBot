import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.dispatcher.initializeUserBlockStates
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import com.github.kotlintelegrambot.entities.keyboard.InlineKeyboardButton
import com.github.kotlintelegrambot.keyboards.Keyboards
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch

object TestBlock {
    fun callTestBlock(chatId: Long, bot: Bot) {
        println("callTestBlock: Инициализация состояний для тестового блока")
        initializeUserBlockStates(chatId, Config.TABLE_FILE)
        val (block1Completed, block2Completed, block3Completed) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
        if (block1Completed && block2Completed && block3Completed) {
            Globals.userBlocks[chatId] = 4
            checkBlocksBeforeTest(chatId, bot, Config.TABLE_FILE)
        } else {
            val notCompletedBlocks = mutableListOf<String>()
            if (!block1Completed) notCompletedBlocks.add("Блок 1")
            if (!block2Completed) notCompletedBlocks.add("Блок 2")
            if (!block3Completed) notCompletedBlocks.add("Блок 3")
            val messageText = "Вы не завершили следующие блоки:\n" +
                    notCompletedBlocks.joinToString("\n") + "\nПройдите их перед тестом."
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    replyMarkup = InlineKeyboardMarkup.createSingleRowKeyboard(
                        InlineKeyboardButton.CallbackData("Вернуться к блокам", "main_menu")
                    )
                )
            }
        }
    }
    fun checkBlocksBeforeTest(chatId: Long, bot: Bot, filePath: String) {
        println("main checkBlocksBeforeTest 277. проверяет, завершены ли все три блока пользователя перед запуском теста. Сначала она инициализирует состояния блоков, затем извлекает их статусы. Если все блоки завершены, вызывается функция для тестового блока (handleBlockTest). В противном случае формируется сообщение с перечнем незавершенных блоков и отправляется пользователю вместе с кнопкой для возврата к блокам.")
        initializeUserBlockStates(chatId, filePath)
        val (block1Completed, block2Completed, block3Completed) = Globals.userBlockCompleted[chatId] ?: Triple(false, false, false)
        if (block1Completed && block2Completed && block3Completed) {
            TestBlock.handleBlockTest(chatId, bot)
        } else {
            val notCompletedBlocks = mutableListOf<String>()
            if (!block1Completed) notCompletedBlocks.add("Блок 1")
            if (!block2Completed) notCompletedBlocks.add("Блок 2")
            if (!block3Completed) notCompletedBlocks.add("Блок 3")
            val messageText = "Вы не выполнили следующие блоки:\n" +
                    notCompletedBlocks.joinToString("\n") +
                    "\nПройдите их перед тестом."
            GlobalScope.launch {
                TelegramMessageService.updateOrSendMessage(
                    chatId = chatId,
                    text = messageText,
                    replyMarkup = Keyboards.returnToBlocksButton()
                )
            }
        }
    }
    fun handleBlockTest(chatId: Long, bot: Bot) {
        println("main handleBlockTest 177. отправляет пользователю сообщение \"Еще не реализовано.\" с клавиатурой для перехода к блоку прилагательных")
        GlobalScope.launch {
            TelegramMessageService.updateOrSendMessage(
                chatId = chatId,
                text = "Еще не реализовано.",
                replyMarkup = Keyboards.goToAdjectivesButton()
            )
        }
        println("main handleBlockTest 178. Сообщение 'Еще не реализовано.' отправлено")
    }
}
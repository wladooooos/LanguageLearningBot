import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import kotlinx.coroutines.suspendCancellableCoroutine
import okhttp3.*
import org.json.JSONObject
import java.io.IOException
import java.net.URLEncoder
import kotlin.coroutines.resume
import kotlin.coroutines.resumeWithException


object TelegramMessageService {
    // Глобальный клиент OkHttp, переиспользуемый во всех вызовах
    val okHttpClient = OkHttpClient()

    // suspend-функция для редактирования сообщения через Telegram API с использованием OkHttp
    suspend fun editMessageTextViaHttpSuspend(
        chatId: Long,
        messageId: Long,
        text: String,
        replyMarkupJson: String? = null
    ): String = suspendCancellableCoroutine { cont ->
        val botToken = Config.BOT_TOKEN
        val url = "https://api.telegram.org/bot$botToken/editMessageText"
        // Формирование тела запроса в формате form-data
        val formBodyBuilder = FormBody.Builder()
            .add("chat_id", chatId.toString())
            .add("message_id", messageId.toString())
            .add("text", text)
        if (replyMarkupJson != null) {
            formBodyBuilder.add("reply_markup", replyMarkupJson)
        }
        val requestBody = formBodyBuilder.build()

        val request = Request.Builder()
            .url(url)
            .post(requestBody)
            .build()

        println("editMessageTextViaHttpSuspend: Отправляем запрос к $url для chatId=$chatId, messageId=$messageId")
        okHttpClient.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                println("editMessageTextViaHttpSuspend: Ошибка запроса: ${e.message}")
                if (cont.isActive) cont.resumeWithException(e)
            }

            override fun onResponse(call: Call, response: Response) {
                try {
                    val responseText = response.body?.string() ?: ""
                    println("editMessageTextViaHttpSuspend: Получен ответ: $responseText")
                    if (cont.isActive) cont.resume(responseText)
                } catch (ex: Exception) {
                    if (cont.isActive) cont.resumeWithException(ex)
                }
            }
        })
    }

    // suspend-функция для отправки нового сообщения через Telegram API с использованием OkHttp
    suspend fun sendMessageViaHttpSuspend(
        chatId: Long,
        text: String,
        replyMarkupJson: String? = null
    ): String = suspendCancellableCoroutine { cont ->
        val botToken = Config.BOT_TOKEN
        val url = "https://api.telegram.org/bot$botToken/sendMessage"
        val formBodyBuilder = FormBody.Builder()
            .add("chat_id", chatId.toString())
            .add("text", text)
        if (replyMarkupJson != null) {
            formBodyBuilder.add("reply_markup", replyMarkupJson)
        }
        val requestBody = formBodyBuilder.build()

        val request = Request.Builder()
            .url(url)
            .post(requestBody)
            .build()

        println("sendMessageViaHttpSuspend: Отправляем запрос к $url для chatId=$chatId с текстом: $text")
        okHttpClient.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                println("sendMessageViaHttpSuspend: Ошибка запроса: ${e.message}")
                if (cont.isActive) cont.resumeWithException(e)
            }

            override fun onResponse(call: Call, response: Response) {
                try {
                    val responseText = response.body?.string() ?: ""
                    println("sendMessageViaHttpSuspend: Получен ответ: $responseText")
                    if (cont.isActive) cont.resume(responseText)
                } catch (ex: Exception) {
                    if (cont.isActive) cont.resumeWithException(ex)
                }
            }
        })
    }

    // Основная suspend-функция, которая пытается отредактировать существующее сообщение,
// а если редактирование не удалось – отправляет новое сообщение.
    suspend fun updateOrSendMessage(
        chatId: Long,
        text: String,
        replyMarkup: InlineKeyboardMarkup? = null
    ) {
        println("updateOrSendMessage: Вызывается для chatId=$chatId с текстом: $text")
        // Предполагаем, что replyMarkup.toString() выдаёт корректный JSON
        val replyMarkupJson = replyMarkup?.toString()
        val currentMessageId = Globals.userMainMessageId[chatId]
        if (currentMessageId != null) {
            println("updateOrSendMessage: Найден сохранённый message_id = $currentMessageId")
            try {
                val response = editMessageTextViaHttpSuspend(chatId, currentMessageId.toLong(), text, replyMarkupJson)
                println("updateOrSendMessage: Редактирование прошло, ответ: $response")
                val json = JSONObject(response)
                if (!json.getBoolean("ok")) {
                    println("updateOrSendMessage: Ответ от Telegram не успешен, отправляем новое сообщение.")
                    Globals.userMainMessageId.remove(chatId)
                } else {
                    return
                }
            } catch (e: Exception) {
                println("updateOrSendMessage: Ошибка редактирования сообщения: ${e.message}")
                println("updateOrSendMessage: Переходим к отправке нового сообщения.")
            }
        } else {
            println("updateOrSendMessage: Сохранённого message_id не найдено, отправляем новое сообщение.")
        }
        // Отправляем новое сообщение через HTTP
        val newResponse = sendMessageViaHttpSuspend(chatId, text, replyMarkupJson)
        println("updateOrSendMessage: Ответ нового сообщения: $newResponse")
        val newJson = JSONObject(newResponse)
        if (!newJson.getBoolean("ok")) {
            throw Exception("Не удалось отправить новое сообщение, ответ: $newResponse")
        }
        val newMessageId = newJson.getJSONObject("result").getInt("message_id")
        Globals.userMainMessageId[chatId] = newMessageId
        println("updateOrSendMessage: Новое сообщение отправлено, сохранён message_id = $newMessageId")
    }

}
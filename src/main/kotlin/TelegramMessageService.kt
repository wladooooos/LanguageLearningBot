import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import kotlinx.coroutines.suspendCancellableCoroutine
import okhttp3.*
import org.json.JSONObject
import java.io.IOException
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
        replyMarkupJson: String? = null,
        parseMode: String? = null // добавили параметр
    ): String = suspendCancellableCoroutine { cont ->
        println("TelegramMessageService editMessageTextViaHttpSuspend")
        val botToken = Config.BOT_TOKEN
        val url = "https://api.telegram.org/bot$botToken/editMessageText"
        println("editMessageTextViaHttpSuspend: наличие parse_mode: ${parseMode}")
        // Формирование тела запроса в формате form-data
        val formBodyBuilder = FormBody.Builder()
            .add("chat_id", chatId.toString())
            .add("message_id", messageId.toString())
            .add("text", text)
            .add("parse_mode", "MarkdownV2")
        if (replyMarkupJson != null) {
            formBodyBuilder.add("reply_markup", replyMarkupJson)
        }
        if (parseMode != null) {
            formBodyBuilder.add("parse_mode", "MarkdownV2")
        }
        val requestBody = formBodyBuilder.build()


// Логируем все параметры из requestBody (если он является FormBody)
        if (requestBody is FormBody) {
            println("TelegramMessageService editMessageTextViaHttpSuspend: Параметры запроса:")
            for (i in 0 until requestBody.size) {
                println("TelegramMessageService editMessageTextViaHttpSuspend: ${requestBody.name(i)} = ${requestBody.value(i)}")
            }
        }

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
        }
        )
    }

    // suspend-функция для отправки нового сообщения через Telegram API с использованием OkHttp
    suspend fun sendMessageViaHttpSuspend(
        chatId: Long,
        text: String,
        replyMarkupJson: String? = null,
        parseMode: String? = null
    ): String = suspendCancellableCoroutine { cont ->
        println("TelegramMessageService ")
        val botToken = Config.BOT_TOKEN
        val url = "https://api.telegram.org/bot$botToken/sendMessage"
        val formBodyBuilder = FormBody.Builder()
            .add("chat_id", chatId.toString())
            .add("text", text)
        if (replyMarkupJson != null) {
            formBodyBuilder.add("reply_markup", replyMarkupJson)
        }
        if (parseMode != null) {
            formBodyBuilder.add("parse_mode", "MarkdownV2")
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

    suspend fun updateOrSendMessage(
        chatId: Long,
        text: String,
        replyMarkup: InlineKeyboardMarkup? = null,
        parseMode: String? = null // опциональный параметр, например, "MarkdownV2"
    ) {
        println("TelegramMessageService updateOrSendMessage: Вызывается для chatId=$chatId с текстом: $text")
        // Добавляем невидимый символ в конец текста, чтобы гарантированно вызвать пересчет и обновление на мобильном клиенте
        val updatedText = text + "\u200B"
        // Предполагаем, что replyMarkup.toString() выдаёт корректный JSON
        val replyMarkupJson = replyMarkup?.toString()
        val currentMessageId = Globals.userMainMessageId[chatId]
        if (currentMessageId != null) {
            println("updateOrSendMessage: Найден сохранённый message_id = $currentMessageId")
            try {
                // Передаём parseMode в запрос редактирования
                val response = editMessageTextViaHttpSuspend(chatId, currentMessageId.toLong(), updatedText, replyMarkupJson, parseMode)
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
        // Передаём parseMode в запрос нового сообщения
        val newResponse = sendMessageViaHttpSuspend(chatId, text, replyMarkupJson, parseMode)
        println("updateOrSendMessage: Ответ нового сообщения: $newResponse")
        val newJson = JSONObject(newResponse)
        if (!newJson.getBoolean("ok")) {
            throw Exception("Не удалось отправить новое сообщение, ответ: $newResponse")
        }
        val newMessageId = newJson.getJSONObject("result").getInt("message_id")
        Globals.userMainMessageId[chatId] = newMessageId
        println("updateOrSendMessage: Новое сообщение отправлено, сохранён message_id = $newMessageId")
    }

    //Ниже - отправка сообщений для блоков без Markdown!!!!!!

    suspend fun editMessageTextViaHttpSuspendWithoutMarkdown(
        chatId: Long,
        messageId: Long,
        text: String,
        replyMarkupJson: String? = null
    ): String = suspendCancellableCoroutine { cont ->
        println("editMessageTextViaHttpSuspendWithoutMarkdown: отправляем запрос для редактирования")
        val botToken = Config.BOT_TOKEN
        val url = "https://api.telegram.org/bot$botToken/editMessageText"
        val formBodyBuilder = FormBody.Builder()
            .add("chat_id", chatId.toString())
            .add("message_id", messageId.toString())
            .add("text", text)
        if (replyMarkupJson != null) {
            formBodyBuilder.add("reply_markup", replyMarkupJson)
        }
        // Не добавляем parse_mode
        val requestBody = formBodyBuilder.build()
        val request = Request.Builder()
            .url(url)
            .post(requestBody)
            .build()

        okHttpClient.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                println("editMessageTextViaHttpSuspendWithoutMarkdown: Ошибка запроса: ${e.message}")
                if (cont.isActive) cont.resumeWithException(e)
            }
            override fun onResponse(call: Call, response: Response) {
                try {
                    val responseText = response.body?.string() ?: ""
                    println("editMessageTextViaHttpSuspendWithoutMarkdown: Получен ответ: $responseText")
                    if (cont.isActive) cont.resume(responseText)
                } catch (ex: Exception) {
                    if (cont.isActive) cont.resumeWithException(ex)
                }
            }
        }
        )
    }

    suspend fun sendMessageViaHttpSuspendWithoutMarkdown(
        chatId: Long,
        text: String,
        replyMarkupJson: String? = null
    ): String = suspendCancellableCoroutine { cont ->
        println("sendMessageViaHttpSuspendWithoutMarkdown: Отправляем запрос для отправки нового сообщения")
        val botToken = Config.BOT_TOKEN
        val url = "https://api.telegram.org/bot$botToken/sendMessage"
        val formBodyBuilder = FormBody.Builder()
            .add("chat_id", chatId.toString())
            .add("text", text)
        if (replyMarkupJson != null) {
            formBodyBuilder.add("reply_markup", replyMarkupJson)
        }
        // Не добавляем parse_mode
        val requestBody = formBodyBuilder.build()
        val request = Request.Builder()
            .url(url)
            .post(requestBody)
            .build()

        okHttpClient.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                println("sendMessageViaHttpSuspendWithoutMarkdown: Ошибка запроса: ${e.message}")
                if (cont.isActive) cont.resumeWithException(e)
            }
            override fun onResponse(call: Call, response: Response) {
                try {
                    val responseText = response.body?.string() ?: ""
                    println("sendMessageViaHttpSuspendWithoutMarkdown: Получен ответ: $responseText")
                    if (cont.isActive) cont.resume(responseText)
                } catch (ex: Exception) {
                    if (cont.isActive) cont.resumeWithException(ex)
                }
            }
        })
    }

    suspend fun updateOrSendMessageWithoutMarkdown(
        chatId: Long,
        text: String,
        replyMarkup: InlineKeyboardMarkup? = null
    ) {
        println("updateOrSendMessageWithoutMarkdown: Вызывается для chatId=$chatId с текстом: $text")
        val updatedText = text + "\u200B"
        val replyMarkupJson = replyMarkup?.toString()
        val currentMessageId = Globals.userMainMessageId[chatId]
        if (currentMessageId != null) {
            println("updateOrSendMessageWithoutMarkdown: Найден сохранённый message_id = $currentMessageId")
            try {
                val response = editMessageTextViaHttpSuspendWithoutMarkdown(chatId, currentMessageId.toLong(), updatedText, replyMarkupJson)
                println("updateOrSendMessageWithoutMarkdown: Редактирование прошло, ответ: $response")
                val json = JSONObject(response)
                if (!json.getBoolean("ok")) {
                    println("updateOrSendMessageWithoutMarkdown: Ответ от Telegram не успешен, отправляем новое сообщение.")
                    Globals.userMainMessageId.remove(chatId)
                } else {
                    return
                }
            } catch (e: Exception) {
                println("updateOrSendMessageWithoutMarkdown: Ошибка редактирования сообщения: ${e.message}")
                println("updateOrSendMessageWithoutMarkdown: Переходим к отправке нового сообщения.")
            }
        } else {
            println("updateOrSendMessageWithoutMarkdown: Сохранённого message_id не найдено, отправляем новое сообщение.")
        }
        val newResponse = sendMessageViaHttpSuspendWithoutMarkdown(chatId, updatedText, replyMarkupJson)
        println("updateOrSendMessageWithoutMarkdown: Ответ нового сообщения: $newResponse")
        val newJson = JSONObject(newResponse)
        if (!newJson.getBoolean("ok")) {
            throw Exception("Не удалось отправить новое сообщение, ответ: $newResponse")
        }
        val newMessageId = newJson.getJSONObject("result").getInt("message_id")
        Globals.userMainMessageId[chatId] = newMessageId
        println("updateOrSendMessageWithoutMarkdown: Новое сообщение отправлено, сохранён message_id = $newMessageId")
    }


}
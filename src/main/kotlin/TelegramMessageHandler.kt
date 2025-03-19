package com.github.kotlintelegrambot.dispatcher

import com.github.kotlintelegrambot.entities.Message
import java.net.HttpURLConnection
import java.net.URL
import java.util.logging.Handler
import java.util.logging.Level
import java.util.logging.LogRecord
import java.util.logging.Logger
import com.github.kotlintelegrambot.Bot
import com.github.kotlintelegrambot.config.Config
import com.github.kotlintelegrambot.entities.ChatId
import com.github.kotlintelegrambot.entities.InlineKeyboardMarkup
import org.json.JSONObject
import java.net.URLEncoder


class TelegramMessageHandler {
    private val botId = 7166917752L

    fun deleteMessageFromChat(chatId: Long) {
        val messageId = TelegramMessageHandler().setupCustomLogger()
        val botToken = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE"
        val url = URL("https://api.telegram.org/bot$botToken/deleteMessage")

        val parameters = "chat_id=$chatId&message_id=$messageId"
        val connection = url.openConnection() as HttpURLConnection
        connection.requestMethod = "POST"
        connection.doOutput = true
        connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded")

        // Отправляем параметры запроса
        connection.outputStream.use { outputStream ->
            outputStream.write(parameters.toByteArray())
        }

        // Проверка успешности удаления
        val responseCode = connection.responseCode
        if (responseCode == 200) {
            println("Сообщение $messageId успешно удалено из чата $chatId")
        } else {
            println("Сообщение $messageId не удалено. Response code: $responseCode")
        }
    }

    fun deleteUserMessageFromChat(chatId: Long, messageId: Long) {
        if (messageId != null) {
            val botToken = "7166917752:AAF0q8pmZBAmgUy_qYUMuR1gJBLaD36VSCE"
            val url = URL("https://api.telegram.org/bot$botToken/deleteMessage")
            val parameters = "chat_id=$chatId&message_id=${messageId}"
            val connection = url.openConnection() as HttpURLConnection
            connection.requestMethod = "POST"
            connection.doOutput = true
            connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded")

            // Отправляем параметры запроса
            connection.outputStream.use { outputStream ->
                outputStream.write(parameters.toByteArray())
            }

            // Проверяем успешность удаления
            val responseCode = connection.responseCode
            if (responseCode == 200) {
                println("Сообщение $messageId успешно удалено из чата $chatId")
            } else {
                println("Сообщение $messageId не удалено. Response code: $responseCode")
            }
        }
    }


    fun setupCustomLogger(): Long? {
        val logger = Logger.getLogger("") // Получаем корневой логгер
        logger.level = Level.INFO // Устанавливаем уровень логирования
        val botId = 7166917752L


        // Очищаем все существующие обработчики
        for (handler in logger.handlers) {
            logger.removeHandler(handler)
        }

        // Добавляем свой обработчик
        logger.addHandler(object : Handler() {
            override fun publish(record: LogRecord) {
                val message = record.message
                if (message.contains("\"ok\":true") && message.contains("\"message_id\"")) {
                    println("Intercepted log: $message")
                }
                val messageId = TelegramMessageHandler().extractMessageIdFromLog(message)
                val botIdFromMessage = TelegramMessageHandler().extractBotIdFromLog(message)
                if (botIdFromMessage == botId) {
                    messageIdList.messageIdList.add(messageId)
                    println("ID последнего сообщения от бота: $messageId")
                }
            }


            override fun flush() {}
            override fun close() {}
        })
        println("ID последнего сообщения от бота (last index): ${messageIdList.messageIdList.lastOrNull()}")
        // Возвращаем последний найденный message_id (если есть)
        return messageIdList.messageIdList.lastOrNull()
    }

    fun extractMessageIdFromLog(logMessage: String): Long? {
        // Используем регулярное выражение или JSON-парсер
        return Regex("\"message_id\":(\\d+)").find(logMessage)?.groupValues?.get(1)?.toLong()
    }

    private fun extractBotIdFromLog(logMessage: String): Long? {
        // Используем регулярное выражение для извлечения bot_id из поля "from": {"id":<bot_id>}
        return Regex("\"from\":\\s*\\{[^}]*\"id\":(\\d+)").find(logMessage)?.groupValues?.get(1)?.toLong()
    }
}

object messageIdList{
    val messageIdList = mutableListOf<Long?>()

    fun addMessageId(messageId: Long) {
        messageIdList.add(messageId)
        println("Добавлено сообщение с ID: $messageId")
    }

    fun getAllMessageIds(): List<Long> {
        return messageIdList.toList() as List<Long> // Возвращаем копию списка для безопасности
    }

    fun getLastMessageId(): Long? {
        return messageIdList.lastOrNull()
    }

    fun clearMessages() {
        messageIdList.clear()
        println("Все сообщения удалены из хранилища.")
    }

    fun sendMessageViaHttp(
        chatId: Long,
        text: String,
        replyMarkup: String? = null  // replyMarkup в виде JSON-строки
    ): Int? {
        println("sendMessageViaHttp: Отправка сообщения через HTTP для chatId=$chatId")
        val botToken = Config.BOT_TOKEN
        val urlString = "https://api.telegram.org/bot$botToken/sendMessage"
        val url = URL(urlString)
        val connection = url.openConnection() as HttpURLConnection
        connection.requestMethod = "POST"
        connection.doOutput = true
        connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded")

        // Формирование параметров запроса
        val params = StringBuilder().apply {
            append("chat_id=$chatId")
            append("&text=" + URLEncoder.encode(text, "UTF-8"))
            if (replyMarkup != null) {
                append("&reply_markup=" + URLEncoder.encode(replyMarkup, "UTF-8"))
            }
        }.toString()
        println("sendMessageViaHttp: Параметры запроса: $params")

        connection.outputStream.use { output ->
            output.write(params.toByteArray())
        }

        // Чтение ответа
        val responseCode = connection.responseCode
        println("sendMessageViaHttp: HTTP Response code: $responseCode")
        if (responseCode == HttpURLConnection.HTTP_OK) {
            val responseText = connection.inputStream.bufferedReader().use { it.readText() }
            println("sendMessageViaHttp: Ответ сервера: $responseText")
            try {
                val json = org.json.JSONObject(responseText)
                if (json.getBoolean("ok")) {
                    val result = json.getJSONObject("result")
                    val messageId = result.getInt("message_id")
                    println("sendMessageViaHttp: Успешно получен message_id = $messageId")
                    return messageId
                } else {
                    println("sendMessageViaHttp: Telegram API вернул ok = false: $responseText")
                }
            } catch (e: Exception) {
                println("sendMessageViaHttp: Ошибка парсинга ответа: ${e.message}")
            }
        } else {
            println("sendMessageViaHttp: HTTP ошибка: $responseCode")
        }
        return null
    }

    fun editMessageTextViaHttp(
        chatId: Long,
        messageId: Long,
        text: String,
        replyMarkupJson: String? = null
    ): String? {
        println("editMessageTextViaHttp: Отправка запроса editMessageText для chatId=$chatId, messageId=$messageId")
        val botToken = Config.BOT_TOKEN
        val urlString = "https://api.telegram.org/bot$botToken/editMessageText"
        val url = URL(urlString)
        val connection = url.openConnection() as HttpURLConnection
        connection.requestMethod = "POST"
        connection.doOutput = true
        connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded")

        // Формируем параметры запроса
        val params = StringBuilder().apply {
            append("chat_id=$chatId")
            append("&message_id=$messageId")
            append("&text=" + URLEncoder.encode(text, "UTF-8"))
            if (replyMarkupJson != null) {
                append("&reply_markup=" + URLEncoder.encode(replyMarkupJson, "UTF-8"))
            }
        }.toString()

        println("editMessageTextViaHttp: Параметры запроса: $params")
        connection.outputStream.use { output ->
            output.write(params.toByteArray())
        }

        val responseCode = connection.responseCode
        println("editMessageTextViaHttp: HTTP Response code: $responseCode")
        val responseText = connection.inputStream.bufferedReader().use { it.readText() }
        println("editMessageTextViaHttp: Полный ответ от Telegram: $responseText")
        return responseText
    }
}
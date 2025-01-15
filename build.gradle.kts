import org.gradle.kotlin.dsl.invoke

plugins {
    kotlin("jvm") version "2.0.20"
    kotlin("plugin.serialization") version "2.0.20"
    id("com.github.johnrengelman.shadow") version "7.1.2"
}

group = "org.example"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
    maven { url = uri("https://jitpack.io") }
}

dependencies {
    // Библиотеки для Telegram
    implementation("com.github.kotlin-telegram-bot.kotlin-telegram-bot:telegram:6.2.0")
    implementation("org.jetbrains.kotlinx:kotlinx-serialization-json:1.6.0")
    implementation("org.jetbrains.kotlin:kotlin-stdlib")

    // Работа с Excel
    implementation("org.apache.poi:poi-ooxml:5.2.3")
    implementation("org.apache.poi:poi:5.2.3")
    // Основной модуль Log4j
    implementation("org.apache.logging.log4j:log4j-core:2.20.0")
    implementation("org.apache.logging.log4j:log4j-api:2.20.0")
    // Дополнительно: поддержка для SLF4J, если используется
    implementation("org.apache.logging.log4j:log4j-slf4j-impl:2.20.0")

    // HTTP и Retrofit
    implementation("com.squareup.retrofit2:retrofit:2.9.0")
    implementation("com.squareup.okhttp3:okhttp:4.10.0")
    implementation("com.squareup.okhttp3:logging-interceptor:4.10.0")

    // Работа с JSON
    implementation("org.json:json:20210307")

    // Тестирование
    testImplementation(kotlin("test"))
}

tasks {
    test {
        useJUnitPlatform()
    }
    shadowJar {
        archiveBaseName.set("telegram-bot")
        archiveVersion.set("1.0")
        archiveClassifier.set("") // Убирает "-all" из имени файла
    }
}

kotlin {
    jvmToolchain(17)
}

tasks.jar {
    manifest {
        attributes(
            "Main-Class" to "org.example.MainKt" // Укажи главный класс с функцией main
        )
    }
}

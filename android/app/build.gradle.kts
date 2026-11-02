import java.util.Properties

plugins {
    id("com.android.application")
    id("com.chaquo.python")
}

// Ключ и пароль лежат вне репозитория. Без них собрать подписанный APK нельзя.
val signingFile = File(System.getProperty("user.home"), "Android/keys/signing.properties")

// Код сайта общий с настольной версией: копируем его из корня репозитория в assets.
// При первом запуске приложение распаковывает его на диск (см. MainActivity).
val webappDir = layout.buildDirectory.dir("generated/webapp")
val syncWebapp by tasks.registering(Sync::class) {
    from("../..") {
        include("app.py", "schema.sql", "templates/**", "static/**")
    }
    into(webappDir.map { it.dir("webapp") })
}
tasks.named("preBuild") { dependsOn(syncWebapp) }

android {
    namespace = "ru.recept.book"
    compileSdk = 35

    defaultConfig {
        applicationId = "ru.recept.book"
        minSdk = 24
        targetSdk = 35
        versionCode = 6
        versionName = "1.5"
        ndk {
            // Для Python 3.12 библиотека lxml есть только под 64-битные устройства.
            abiFilters += listOf("arm64-v8a", "x86_64")
        }
    }

    if (signingFile.exists()) {
        val props = Properties().apply { signingFile.inputStream().use { load(it) } }
        signingConfigs {
            create("release") {
                storeFile = file(props.getProperty("storeFile"))
                storePassword = props.getProperty("storePassword")
                keyAlias = props.getProperty("keyAlias")
                keyPassword = props.getProperty("keyPassword")
            }
        }
        buildTypes["release"].signingConfig = signingConfigs["release"]
    }

    sourceSets["main"].assets.srcDir(webappDir)

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
}

dependencies {
    // Нужен только FileProvider: через него камера и другие приложения получают доступ к файлам.
    implementation("androidx.core:core:1.13.1")
}

chaquopy {
    defaultConfig {
        version = "3.12"
        buildPython("/usr/bin/python3")
        pip {
            install("-r", "../../requirements.txt")
        }
    }
}

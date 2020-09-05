dependencies {
    api(project(":kotlin-xls"))
    api("io.github.microutils:kotlin-logging:1.8.3")
    api("com.sksamuel.hoplite:hoplite-hocon:1.3.5")
    implementation("org.jetbrains.kotlin:kotlin-reflect:1.4.0")
    runtimeOnly("ch.qos.logback:logback-classic:1.2.3")
    testImplementation("com.h2database:h2:1.4.200")
}

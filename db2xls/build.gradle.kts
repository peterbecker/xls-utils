dependencies {
    implementation(project(":kotlin-xls"))
    implementation("io.github.microutils:kotlin-logging:1.8.3")
    implementation("com.sksamuel.hoplite:hoplite-hocon:1.3.5")
    runtimeOnly("ch.qos.logback:logback-classic:1.2.3")
    testImplementation("com.h2database:h2:1.4.200")
}

dependencies {
    implementation(project(":kotlin-xls"))
    implementation("io.github.microutils:kotlin-logging:1.8.3")
    testImplementation("com.h2database:h2:1.4.200")
    testRuntimeOnly("ch.qos.logback:logback-classic:1.2.3")
}

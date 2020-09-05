import org.gradle.jvm.tasks.Jar

dependencies {
    api(project(":db2xls"))
    api("ch.qos.logback:logback-classic:1.2.3")
    api("com.h2database:h2:1.4.200")
    api("org.postgresql:postgresql:42.2.16")
    api("mysql:mysql-connector-java:8.0.21")
    api("com.microsoft.sqlserver:mssql-jdbc:8.4.1.jre11")
}

val fatJar = task("fatJar", type = Jar::class) {
    manifest {
        attributes["Implementation-Title"] = "DB2XLS Bundle"
        attributes["Implementation-Version"] = "0.1-SNAPSHOT"
        attributes["Main-Class"] = "de.peterbecker.xls.Db2XlsKt"
    }
    from(
        configurations.compile.get().map { if (it.isDirectory) it else zipTree(it) },
        configurations.runtimeClasspath.get().map { if (it.isDirectory) it else zipTree(it) }
    )
    with(tasks.jar.get() as CopySpec)
}

tasks {
    "build" {
        dependsOn(fatJar)
    }
}
import org.jetbrains.kotlin.gradle.plugin.KotlinPluginWrapper
import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

group = "de.peterbecker.xls"
version = "1.0-SNAPSHOT"

plugins {
    kotlin("jvm") version "1.3.11"
}

repositories {
    mavenCentral()
}

tasks.withType<KotlinCompile> {
    kotlinOptions.jvmTarget = "1.8"
}

tasks.withType<Wrapper> {
    gradleVersion = "5.3"
    distributionType = Wrapper.DistributionType.ALL
}

subprojects {
    apply<KotlinPluginWrapper>()
    apply<JavaPlugin>()

    repositories {
        mavenCentral()
    }

    dependencies {
        compile(kotlin("stdlib-jdk8"))
        testCompile("junit:junit:4.12")
        testCompile("org.assertj:assertj-core:3.12.1")
    }
}
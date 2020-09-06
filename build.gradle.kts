import org.jetbrains.kotlin.gradle.plugin.KotlinPluginWrapper
import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

group = "de.peterbecker.xls"
version = "1.0-SNAPSHOT"

plugins {
    kotlin("jvm") version "1.4.0"
}

repositories {
    mavenCentral()
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

    tasks.withType<KotlinCompile> {
        kotlinOptions.jvmTarget = "11"
    }

    dependencies {
        implementation(kotlin("stdlib-jdk8"))
        testImplementation("org.junit.jupiter:junit-jupiter:5.6.2")
        testImplementation("org.assertj:assertj-core:3.17.2")
    }
}
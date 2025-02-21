plugins {
    kotlin("jvm") version "2.1.10"
}

group = "org.example"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    testImplementation(kotlin("test"))
    implementation("org.apache.poi:poi:5.2.3")
    implementation("org.apache.poi:poi-ooxml:5.2.3")
    implementation("org.json:json:20231013")
    implementation("org.apache.logging.log4j:log4j-core:2.24.3")

}

tasks.test {
    useJUnitPlatform()
}

kotlin {
    jvmToolchain(20)
}
/*
 * Copyright 2026 Ritense BV, the Netherlands.
 *
 * Licensed under EUPL, Version 1.2 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://joinup.ec.europa.eu/collection/eupl/eupl-text-eupl-12
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

dockerCompose {
    setProjectName("sharepoint-plugin")
    isRequiredBy(project.tasks.test)

    tasks.test {
        useComposeFiles.addAll("$rootDir/docker-resources/docker-compose-base-test.yml", "docker-compose-override.yml")
    }
}

val kotlinLoggingVersion: String by project
val mockitoKotlinVersion: String by project
val valtimoVersion: String by project
val operatonVersion: String by project

dependencies {
    compileOnly("com.ritense.valtimo:plugin-valtimo")
    compileOnly("com.ritense.valtimo:case")
    compileOnly("com.ritense.valtimo:process-document")
    compileOnly("com.ritense.valtimo:zaken-api")
    compileOnly("com.ritense.valtimo:value-resolver")
    compileOnly("com.ritense.valtimo:contract")
    compileOnly("org.operaton.bpm:operaton-engine:$operatonVersion")
    compileOnly("org.springframework.boot:spring-boot-autoconfigure")
    compileOnly("org.springframework.boot:spring-boot-starter-web")
    compileOnly("org.springframework.boot:spring-boot-starter-security")

    implementation("com.microsoft.graph:microsoft-graph:6.65.0")
    implementation("com.azure:azure-identity:1.18.4")

    implementation("org.operaton.spin:operaton-spin-dataformat-json-jackson:1.0.3")
    implementation("org.operaton.bpm:operaton-engine-plugin-spin:1.0.3")

    compileOnly("io.github.oshai:kotlin-logging:$kotlinLoggingVersion")

    // Testing
    testImplementation("com.ritense.valtimo:plugin-valtimo")
    testImplementation("com.ritense.valtimo:process-document")
    testImplementation("com.ritense.valtimo:building-block")
    testImplementation("com.ritense.valtimo:local-resource")
    testImplementation("com.ritense.valtimo:test-utils-common")

    testImplementation("org.springframework.boot:spring-boot-starter-web")
    testImplementation("org.springframework.boot:spring-boot-starter-data-jpa")
    testImplementation("org.springframework.boot:spring-boot-starter-security")
    testImplementation("org.springframework.boot:spring-boot-starter-test")

    testImplementation("org.mockito:mockito-core")
    testImplementation("org.mockito.kotlin:mockito-kotlin:$mockitoKotlinVersion")

    testImplementation("org.postgresql:postgresql")

    testImplementation("org.jetbrains.kotlin:kotlin-test")
    testImplementation("org.jetbrains.kotlin:kotlin-test-junit5")
}

apply(from = "gradle/publishing.gradle")

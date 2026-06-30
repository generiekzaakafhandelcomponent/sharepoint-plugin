val kotlinLoggingVersion: String by project
val nettyResolverDnsNativeMacOsVersion: String by project

val valtimoVersion: String by project

dependencies {
    implementation(platform("com.ritense.valtimo:valtimo-dependency-versions:$valtimoVersion"))

    implementation("com.ritense.valtimo:valtimo-dependencies:$valtimoVersion")
    implementation("com.ritense.valtimo:valtimo-gzac-dependencies:$valtimoVersion")
    implementation("com.ritense.valtimo:local-mail:$valtimoVersion")

    implementation("org.springframework.boot:spring-boot-starter-web")
    implementation("org.postgresql:postgresql")
    implementation("io.github.oshai:kotlin-logging:$kotlinLoggingVersion")

    if (System.getProperty("os.arch") == "aarch64") {
        runtimeOnly("io.netty:netty-resolver-dns-native-macos:$nettyResolverDnsNativeMacOsVersion:osx-aarch_64")
    }

    implementation(project(":backend:plugin"))
}

tasks.jar {
    enabled = false
}

apply(from = "../../gradle/environment.gradle.kts")
val configureEnvironment = extra["configureEnvironment"] as (task: ProcessForkOptions) -> Unit

dockerCompose {
    setProjectName("gzac-docker-compose")
    composeAdditionalArgs.addAll("--profile", "zgw")
    // composeAdditionalArgs = ["--profile", "zgw", "--profile", "openformulieren", "--profile", "openklant"]
    stopContainers = false
    removeContainers = false
    removeVolumes = false
}

tasks.bootRun {
    dependsOn("composeUp")
    systemProperty("spring.profiles.include", "dev")
    val t = this
    doFirst {
        configureEnvironment(t)
    }
}

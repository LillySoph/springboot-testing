import com.github.gradle.node.npm.task.NpmTask

plugins {
	java
	id("org.springframework.boot") version "3.2.0"
	id("io.spring.dependency-management") version "1.1.4"
	id("com.github.node-gradle.node") version "7.0.1"
}

group = "lilly.testing"
version = "0.0.1-SNAPSHOT"

java {
	sourceCompatibility = JavaVersion.VERSION_17
}

configurations {
	compileOnly {
		extendsFrom(configurations.annotationProcessor.get())
	}
}

repositories {
	mavenCentral()
}

dependencies {
	implementation("org.springframework.boot:spring-boot-starter-data-jpa")
	implementation("org.springframework.boot:spring-boot-starter-oauth2-client")
	implementation("org.springframework.boot:spring-boot-starter-thymeleaf")
	// apache poi
	implementation("org.apache.poi:poi:5.2.3")
	implementation("org.apache.poi:poi-ooxml:5.2.3")
	implementation("org.apache.poi:poi-scratchpad:5.2.3")
	// https://mvnrepository.com/artifact/com.googlecode.json-simple/json-simple
	implementation("com.googlecode.json-simple:json-simple:1.1.1")
	// https://mvnrepository.com/artifact/fr.opensagres.xdocreport/fr.opensagres.poi.xwpf.converter.pdf
	implementation("fr.opensagres.xdocreport:fr.opensagres.poi.xwpf.converter.core:2.0.4")
	implementation("fr.opensagres.xdocreport:fr.opensagres.poi.xwpf.converter.pdf:2.0.4")
	// https://mvnrepository.com/artifact/com.itextpdf/itextpdf
	implementation("com.itextpdf:itextpdf:5.5.13.3")
	compileOnly("org.projectlombok:lombok")
	runtimeOnly("com.h2database:h2")
	annotationProcessor("org.projectlombok:lombok")
	testImplementation("org.springframework.boot:spring-boot-starter-test")
}

// node and npm
node {
	download.set(true)
	version.set("20.10.0")
	npmInstallCommand.set("ci")
	nodeProjectDir.set(file("$projectDir/../VueTesting"))
}

// frontend
tasks.register<NpmTask>("buildWebApp") {
	dependsOn("npmInstall")
	args.set(
			listOf(
					"run",
					"build"
			)
	)
}

tasks.register<Copy>("copyWebApp") {
	from("$projectDir/../VueTesting/dist")
	into(layout.buildDirectory.dir("resources/main/static"))
}

// java plugin
tasks.compileTestJava {
	mustRunAfter("copyWebApp")
}

tasks.compileJava {
	mustRunAfter("copyWebApp")
}

tasks.jar {
	archiveBaseName.set(rootProject.name)
}

// tests
tasks.withType<Test> {
	useJUnitPlatform()
}

tasks.bootRun {
	dependsOn("clean")
	dependsOn("buildWebApp")
	dependsOn("copyWebApp")
	mustRunAfter("clean")
}

// packaging
tasks.bootJar {
	dependsOn("clean")
	dependsOn("buildWebApp")
	dependsOn("copyWebApp")
	mustRunAfter("clean")
}

// todo: deploy

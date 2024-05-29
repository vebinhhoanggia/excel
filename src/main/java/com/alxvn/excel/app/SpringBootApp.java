package com.alxvn.excel.app;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javafx.application.Application;

@SpringBootApplication
public class SpringBootApp {

	public static void main(String[] args) {
		// Start the Spring Boot application
		SpringApplication.run(SpringBootApp.class, args);

		// Launch the JavaFX application
		Application.launch(JavaFxApplication.class, args);
	}
}

package com.alxvn.excel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;

import com.alxvn.excel.app.JavaFxApplication;

import javafx.application.Application;

@SpringBootApplication
public class SpringBootApp extends SpringBootServletInitializer {
	public static void main(final String[] args) {
		SpringApplication.run(SpringBootApp.class, args);
		// Launch the JavaFX application
		Application.launch(JavaFxApplication.class, args);
	}
}
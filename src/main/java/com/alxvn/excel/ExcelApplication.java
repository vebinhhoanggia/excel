package com.alxvn.excel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;

@SpringBootApplication
public class ExcelApplication extends SpringBootServletInitializer {

	public static void main(final String[] args) {
		SpringApplication.run(ExcelApplication.class, args);
	}

}

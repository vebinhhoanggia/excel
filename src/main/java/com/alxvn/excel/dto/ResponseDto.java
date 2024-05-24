package com.alxvn.excel.dto;

import org.springframework.core.io.InputStreamResource;

public class ResponseDto {
	
	private final InputStreamResource file;
	private final String message;

	public ResponseDto(InputStreamResource file, String message) {
		this.file = file;
		this.message = message;
	}

	public InputStreamResource getFile() {
		return file;
	}

	public String getMessage() {
		return message;
	}
}

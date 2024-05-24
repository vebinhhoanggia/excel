package com.alxvn.excel.dto;

import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;

public class DownloadResponse {
	private final InputStreamResource file;
	private final String fileName;
	private final String contentType;
	private final String message;

	public DownloadResponse(InputStreamResource file, String fileName, String contentType, String message) {
		this.file = file;
		this.fileName = fileName;
		this.contentType = contentType;
		this.message = message;
	}

	public InputStreamResource getFile() {
		return file;
	}

	public String getFileName() {
		return fileName;
	}

	public String getContentType() {
		return contentType;
	}

	public String getMessage() {
		return message;
	}

	public HttpHeaders getHttpHeaders() {
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.parseMediaType(contentType));
		headers.setContentDisposition(
				org.springframework.http.ContentDisposition.attachment().filename(fileName).build());
		return headers;
	}
}

/**
 * 
 */
package com.alxvn.excel.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Deque;
import java.util.LinkedList;
import java.util.List;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;

import com.alxvn.excel.dto.DownloadResponse;
import com.alxvn.excel.dto.ResponseDto;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 
 */
public class FileUtil {
	
	
	private final static ObjectMapper objectMapper = new ObjectMapper();
	
	private static final Logger log = LoggerFactory.getLogger(FileUtil.class);
	
	public  void zipFolderByPath(final String folderPath, final String outputZipFile) {
		try {
			final List<Path> filesToZip = new ArrayList<>();
			try (var stream = Files.walk(Paths.get(folderPath))) {
				stream.filter(Files::isRegularFile).forEach(filesToZip::add);
			}

			try (var zipOutputStream = new ZipOutputStream(new FileOutputStream(outputZipFile))) {
				for (final Path filePath : filesToZip) {
					final var entryName = folderPath.substring(folderPath.lastIndexOf("/") + 1) + "/"
							+ filePath.getFileName().toString();
					zipOutputStream.putNextEntry(new ZipEntry(entryName));

					try (var fileInputStream = new FileInputStream(filePath.toFile())) {
						int read;
						while ((read = fileInputStream.read()) != -1) {
							zipOutputStream.write(read);
						}
					}
					zipOutputStream.closeEntry();
				}
			}
		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

	public  String getLastFolderName(final String folderPath) {
		// Tách folder path thành các phần tử
		final var folders = folderPath.split(Pattern.quote(File.separator));

		// Lấy tên thư mục cuối cùng
		final var lastFolderName = folders[folders.length - 1];

		return lastFolderName;
	}

	public  ResponseEntity<InputStreamResource> zipFolder(final String folderPath) {
		log.debug("Bắt đầu xử lý tạo zip file !!!");
		if (StringUtils.isBlank(folderPath)) {
			return ResponseEntity.noContent().build();
		}

		try (var byteArrayOutputStream = new ByteArrayOutputStream()) {
			zipFolder(folderPath, byteArrayOutputStream);
			final var inputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
			final var resource = new InputStreamResource(inputStream);

			// Construct the new file name
			final var fileDownloadName = getLastFolderName(folderPath) + ".zip";

//			final var headers = new HttpHeaders();
//			headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
//			headers.setContentDisposition(ContentDisposition.builder("attachment").filename(fileDownloadName).build());

			String contentType = "application/zip";

			// Create the DownloadResponse object
			DownloadResponse downloadResponse = new DownloadResponse(resource, fileDownloadName, contentType, null);

			log.debug("Kết thúc xử lý zip file !!!");
			// Return the response with the appropriate headers
			return ResponseEntity.ok().headers(downloadResponse.getHttpHeaders()).body(downloadResponse.getFile());

//			return new ResponseEntity<>(resource, headers, HttpStatus.OK);
		} catch (final IOException e) {
			// Handle exceptions here
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}

	public  void zipFolder(final String folderPath, final OutputStream outputStream) throws IOException {
		final var path = Paths.get(folderPath);
		try (var zipOutputStream = new ZipOutputStream(outputStream)) {
			addFolderToZip(zipOutputStream, path, "");
		}
	}

	private  void addFolderToZip(final ZipOutputStream zipOutputStream, final Path path, final String basePath)
			throws IOException {
		final Deque<Path> directories = new LinkedList<>();
		directories.offerFirst(path);

		while (!directories.isEmpty()) {
			final var currentPath = directories.pollFirst();
			final var relativePath = basePath + path.relativize(currentPath).toString();

			if (Files.isDirectory(currentPath)) {
				zipOutputStream.putNextEntry(new ZipEntry(relativePath + "/"));
				zipOutputStream.closeEntry();

				try (var dirStream = Files.list(currentPath)) {
					dirStream.forEach(directories::offerLast);
				}
			} else {
				zipOutputStream.putNextEntry(new ZipEntry(relativePath));
				Files.copy(currentPath, zipOutputStream);
				zipOutputStream.closeEntry();
			}
		}
	}
}

package com.alxvn.excel.home;

import java.io.IOException;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.alxvn.excel.config.ApiVersion;
import com.alxvn.excel.service.ExcelOperationService;


@RestController
@RequestMapping(value = ExcelOperationController.REQUEST_MAPPING)
public class ExcelOperationController {

	static final String RESOURCE_SCHEMA = "/opeation/excel";

	static final String REQUEST_MAPPING = ApiVersion.V1 + RESOURCE_SCHEMA;

	private static final Logger log = LoggerFactory.getLogger(ExcelOperationController.class);

	@Autowired
	private ExcelOperationService excelService;

	@PostMapping(value = "/splitSheet")
	public ResponseEntity<String> hello(@RequestParam("file") List<MultipartFile> files) {
		log.debug("splitFile:  {}", files);

		excelService.splitSheetExcel(files);

		return ResponseEntity.ok("Files uploaded successfully");
	}

	@PostMapping("/upload-excel")
	public ResponseEntity<InputStreamResource> uploadAndSplitExcelFiles(@RequestParam("files") List<MultipartFile> files)
			throws IOException {
		final double perSheetInFile = 10.0;
		return excelService.uploadAndSplitExcelFiles(files, perSheetInFile);
		
	}

}
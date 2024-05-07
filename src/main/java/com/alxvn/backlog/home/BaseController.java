/**
 *
 */
package com.alxvn.backlog.home;

import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.alxvn.backlog.BacklogService;
import com.alxvn.backlog.config.ApiVersion;
import com.alxvn.backlog.handle.IncorrectFullNameException;
import com.alxvn.backlog.util.BacklogExcelUtil;
import com.opencsv.exceptions.CsvException;

/**
 *
 */
@Controller
@RequestMapping(value = BaseController.REQUEST_MAPPING)
public class BaseController {

	private static final String RESOURCE_SCHEMA = "/base";
	protected static final String REQUEST_MAPPING = ApiVersion.V1 + RESOURCE_SCHEMA;

	@Autowired
	private BacklogService backlogService;

	@GetMapping("genSchedule")
	public String stastics() {
		final var util = new BacklogExcelUtil();
		final var wrPath = "templates/pjjyuji_data_csv_20240415.csv";
		final var backlogPath = "templates/Backlog-Issues-20240415-1157.csv";
		util.createScheduleFromBacklog(wrPath, backlogPath);
		return "index";
	}

	@PostMapping(value = "genSchedule")
	public ResponseEntity<String> genSchedule(@RequestParam("file1") final MultipartFile file1,
			@RequestParam("file2") final MultipartFile file2)
			throws IOException, CsvException, IncorrectFullNameException {

//		final var util = new BacklogExcelUtil();
//		final var wrPath = "templates/pjjyuji_data_csv_20240415.csv";
//		final var backlogPath = "templates/Backlog-Issues-20240415-1157.csv";
//		util.createScheduleFromBacklog(wrPath, backlogPath);

		backlogService.stastics(file1, file2, null);

		return ResponseEntity.ok("Files uploaded successfully");
	}

}
/**
 *
 */
package com.alxvn.backlog.home;

import java.io.IOException;
import java.util.Collection;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.alxvn.backlog.BacklogService;
import com.alxvn.backlog.config.ApiVersion;
import com.alxvn.backlog.handle.IncorrectFullNameException;
import com.alxvn.backlog.schedule.BacklogExcel;
import com.opencsv.exceptions.CsvException;

/**
 *
 */
@RestController
@RequestMapping(value = BaseController.REQUEST_MAPPING)
public class BaseController {

	private static final String RESOURCE_SCHEMA = "/base";
	protected static final String REQUEST_MAPPING = ApiVersion.V1 + RESOURCE_SCHEMA;

	@Autowired
	private BacklogService backlogService;

	@GetMapping("/hello")
	@ResponseBody
	public Collection<String> sayHello() {
		return IntStream.range(0, 10).mapToObj(i -> "Hello number " + i).collect(Collectors.toList());
	}

	@GetMapping("genSchedule")
	public String stastics() {
		final var util = new BacklogExcel();
		final var wrPath = "templates/pjjyuji_data_csv_20240415.csv";
		final var backlogPath = "templates/Backlog-Issues-20240415-1157.csv";
		util.createScheduleFromBacklog(wrPath, backlogPath);
		return "index";
	}

	@PostMapping(value = "genSchedule")
	public ResponseEntity<InputStreamResource> genSchedule(@RequestParam("file1") final MultipartFile file1,
			@RequestParam("file2") final MultipartFile file2)
			throws IOException, CsvException, IncorrectFullNameException {

//		final var util = new BacklogExcelUtil();
//		final var wrPath = "templates/pjjyuji_data_csv_20240415.csv";
//		final var backlogPath = "templates/Backlog-Issues-20240415-1157.csv";
//		util.createScheduleFromBacklog(wrPath, backlogPath);

		final var rootFldResult = backlogService.stastics(file1, file2);

		return backlogService.zipFolder(rootFldResult);

//		return ResponseEntity.ok("create schedule successfully");
	}

}
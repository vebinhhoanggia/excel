/**
 *
 */
package com.alxvn.backlog.home;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import com.alxvn.backlog.schedule.BacklogExcel;

/**
 *
 */
@Controller
public class HomeController {

	@GetMapping("/")
	public String index() {
		return "index";
	}

	@GetMapping("stastics")
	public String stastics() {
		var util = new BacklogExcel();
		final var wrPath = "templates/pjjyuji_data_csv_20240415.csv";
		final var backlogPath = "templates/Backlog-Issues-20240415-1157.csv";
		util.createScheduleFromBacklog(wrPath, backlogPath);
		return "index";
	}
}
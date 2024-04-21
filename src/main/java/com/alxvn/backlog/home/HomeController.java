/**
 *
 */
package com.alxvn.backlog.home;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import com.alxvn.backlog.util.BacklogExcelUtil;

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
		BacklogExcelUtil.createScheduleFromBacklog();
		return "index";
	}
}
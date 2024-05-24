/**
 *
 */
package com.alxvn.excel.home;

import java.util.Collection;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;


/**
 *
 */
@Controller
public class HomeController {

	@GetMapping("/")
	public String index() {
		return "index";
	}

	@GetMapping("/hello")
	@ResponseBody
	public Collection<String> sayHello() {
		return IntStream.range(0, 10).mapToObj(i -> "Hello number " + i).collect(Collectors.toList());
	}

}
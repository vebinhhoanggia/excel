/**
 *
 */
package com.alxvn.excel.home;

import java.io.IOException;
import java.util.Collection;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import com.alxvn.excel.config.ApiVersion;
import com.alxvn.excel.service.ExcelOperationService;

/**
 *
 */
@RestController
@RequestMapping(value = BaseController.REQUEST_MAPPING)
public class BaseController {

	private static final String RESOURCE_SCHEMA = "/base";
	protected static final String REQUEST_MAPPING = ApiVersion.V1 + RESOURCE_SCHEMA;


	@GetMapping("/hello")
	@ResponseBody
	public Collection<String> sayHello() {
		return IntStream.range(0, 10).mapToObj(i -> "Hello number " + i).collect(Collectors.toList());
	}
	
	@Autowired
	private ExcelOperationService excelService;

	@GetMapping("/countEcTc")
	@ResponseBody
	public int countAllTC() throws IOException {
		return excelService.countNumericMergedCells("\\\\192.168.10.40\\Training\\Enercom\\Ngantl\\テストチェックリスト");
	}

}
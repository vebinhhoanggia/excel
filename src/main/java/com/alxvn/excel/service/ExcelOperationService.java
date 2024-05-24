/**
 *
 */
package com.alxvn.excel.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.alxvn.excel.util.FileUtil;
import com.alxvn.excel.util.ScheduleHelper;

/**
 * @author KEDD
 *
 */
@Service
public class ExcelOperationService {

	private static final Logger log = LoggerFactory.getLogger(ExcelOperationService.class);

	public void splitSheetExcel(List<MultipartFile> files) {

	}

	private static final int columnSIndex = CellReference.convertColStringToIndex("S");
	private static final int rowChkIdx = 3;

	private static int getSuffixNumber(File file) {
		final String fileName = file.getName();
		final String pattern = "_suff_(\\d+)";
		final Pattern regex = Pattern.compile(pattern);

		final Matcher matcher1 = regex.matcher(fileName);
		if (matcher1.find()) {
			final String suffix = matcher1.group(1);
			return Integer.parseInt(suffix);
		}
		return 0;
	}

	public String getFileName(MultipartFile file) {
		final String originalFileName = file.getOriginalFilename();

		if (org.springframework.util.StringUtils.hasText(originalFileName)) {
			try {
				// Decode file name using UTF-8 encoding
				return new String(originalFileName.getBytes(StandardCharsets.ISO_8859_1), StandardCharsets.UTF_8);
			} catch (final Exception e) {
				// Handle the exception or return the original file name
				e.printStackTrace();
			}
		}

		return originalFileName;
	}

	public ResponseEntity<Object> uploadAndSplitExcelFiles(@RequestParam("files") List<MultipartFile> files,
			double perSheetInFile) throws IOException {
		System.out.println("Bắt đầu xử lý");
		log.debug("Bắt đầu xử lý");
		final String pathStr = "D:\\Doc\\split\\result";
		List<String> errors = new ArrayList<>();

		final Path folder = Paths.get(pathStr);

		if (Files.exists(folder)) {
			Files.walkFileTree(folder, new SimpleFileVisitor<Path>() {
				@Override
				public FileVisitResult visitFile(final Path file, final BasicFileAttributes attrs) throws IOException {
					Files.delete(file);
					return FileVisitResult.CONTINUE;
				}

				@Override
				public FileVisitResult postVisitDirectory(final Path dir, final IOException exc) throws IOException {
					Files.delete(dir);
					return FileVisitResult.CONTINUE;
				}
			});
			System.out.println("All files and directories within the folder have been deleted.");
		} else {
			Files.createDirectories(folder);
			System.out.println("New folder created: " + folder);
		}

		final Path folderPath = Paths.get(pathStr);
		if (!Files.exists(folderPath)) {
			Files.createDirectories(folderPath);
		}
		try {
			// Iterate over each uploaded file
			for (final MultipartFile file : files) {
				final List<String> allTargetFilePaths = new ArrayList<>();
				List<Pair<String, Double>> preChkIdList = new ArrayList<>();
				boolean isFirst = true;
				final String fileName = getFileName(file);
				// Read the Excel file
				try (final Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
					final int sheetCount = workbook.getNumberOfSheets();
					final double result = sheetCount / perSheetInFile;
					final int count = (int) Math.ceil(result);

					final List<String> targetFilePaths = generateCountSuffixList(fileName, count);
					allTargetFilePaths.addAll(targetFilePaths);
					for (final String targetFilePath : targetFilePaths) {
						final Path filePath = folderPath.resolve(targetFilePath);
						final File tempFile = filePath.toFile();
						FileCopyUtils.copy(file.getBytes(), tempFile);
					}
				}
				for (final String targetFileName : allTargetFilePaths) {
					final Path filePath = folderPath.resolve(targetFileName);
					final File splitFile = filePath.toFile();
					// Read the Excel file
					final List<Pair<String, Double>> curChkIdList = new ArrayList<>();
					try (final Workbook workbook = new XSSFWorkbook(new FileInputStream(splitFile))) {
						final FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
						final int sheetCount = workbook.getNumberOfSheets();
						final int sheetIndexToRemove = getSuffixNumber(splitFile);
						final int startIdx = (sheetIndexToRemove - 1) * (int) perSheetInFile;
						final int endIdx = startIdx + 9;
						final List<Integer> removeAbles = new ArrayList<>();
						for (int i = sheetCount - 1; i >= 0; i--) {
							if (i < startIdx || i > endIdx) {
								removeAbles.add(i);
							}
						}
						removeAbles.forEach(workbook::removeSheetAt);
						final Iterator<Sheet> sheetIterator = workbook.sheetIterator();
						while (sheetIterator.hasNext()) {
							final Sheet sheet = sheetIterator.next();
							final String sheetName = sheet.getSheetName();
							if (!StringUtils.equals(sheetName, "変更履歴")) {
								final Row rCharChkId = sheet.getRow(rowChkIdx);
								final Row rCheckOrdId = sheet.getRow(rowChkIdx + 1);
								for (int column = columnSIndex; column <= rCharChkId.getLastCellNum(); column++) {
									final Cell charCell = rCharChkId.getCell(column); // Get the source cell
									final Cell ordCell = rCheckOrdId.getCell(column); // Get the source cell

									final CellValue charCellValue = formulaEvaluator.evaluate(charCell);
									final String charVal = StringUtils.defaultString(
											ScheduleHelper.getCellValueAsString(charCell, charCellValue),
											ScheduleHelper.getCellValueAsString(charCell));

									final CellValue ordCellValue = formulaEvaluator.evaluate(ordCell);
									final String ordVal = StringUtils.defaultString(
											ScheduleHelper.getCellValueAsString(ordCell, ordCellValue),
											ScheduleHelper.getCellValueAsString(ordCell));
									if (StringUtils.isNotBlank(charVal) || StringUtils.isNotBlank(ordVal)) {
										curChkIdList.add(Pair.of(charVal, Double.valueOf(ordVal)));
									}
								}
							}
						}

						try (FileOutputStream outputStream = new FileOutputStream(splitFile, false)) {
							workbook.write(outputStream);
						}

						final Pair<String, Double> start = curChkIdList.get(0);
						final String startString = start.getKey() + String.format("%03d", start.getValue().intValue());
						final Pair<String, Double> end = curChkIdList.get(curChkIdList.size() - 1);
						final String endString = end.getKey() + String.format("%03d", end.getValue().intValue());
						final String newSuffix = startString + "-" + endString;
						// Use regular expression to replace the suffix
						final String newFileName = targetFileName.replaceAll("_suff_\\d+", "_" + newSuffix);

						final Path targetPath = filePath.resolveSibling(newFileName);

						Files.copy(filePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
						Files.delete(filePath);
						if (isNotIncreasing(curChkIdList) || !isFirst && isNotIncreasing(curChkIdList, preChkIdList)) {
							String str1 = "ListCheckId from:" + newFileName + " : " + curChkIdList;
							String str2 = "File: " + newFileName + " co checkId khong dung thu tu";
							System.out.println(str1);
							System.out.println(str2);
							log.debug(str1);
							log.debug(str2);
							errors.add(str1);
							errors.add(str2);
						}
						preChkIdList = curChkIdList;
						isFirst = false;
					}
				}
			}

			// Create a ZIP file containing the split files
			System.out.println("Kết thúc xử lý");
			log.debug("Kết thúc xử lý");
			
			FileUtil util = new FileUtil();
			return util.zipFolder(pathStr, "");
		} catch (final Exception e) {
			System.out.println("Xử lý lỗi");
			log.debug("Xử lý lỗi");
			e.printStackTrace();
			// Handle the exception appropriately
			return ResponseEntity.badRequest().build();
		}
	}
	

	private static int compareExcelColumnStrings(String str1, String str2) {
		final int len1 = str1.length();
		final int len2 = str2.length();

		// Compare lengths
		if (len1 < len2) {
			return -1;
		}
		if (len1 > len2) {
			return 1;
		}

//		// Compare characters
//		for (int i = 0; i < len1; i++) {
//			final char char1 = str1.charAt(i);
//			final char char2 = str2.charAt(i);
//
//			if (char1 < char2) {
//				return 1;
//			}
//			if (char1 > char2) {
//				return -1;
//			}
//		}

		// Strings are equal
		return str1.compareTo(str2);
	}

	private static int comparePairs(Pair<String, Double> pair1, Pair<String, Double> pair2) {
		int compareResult = compareExcelColumnStrings(pair1.getKey(), pair2.getKey());
		if (compareResult == 0) {
			compareResult = pair1.getValue().compareTo(pair2.getValue());
		}
		return compareResult;
	}

	public static boolean isNotIncreasing(List<Pair<String, Double>> list) {
		for (int i = 1; i < list.size(); i++) {
			final Pair<String, Double> previous = list.get(i - 1);
			final Pair<String, Double> current = list.get(i);

			final int compareResult = comparePairs(current, previous);
			if (compareResult <= 0) {
				System.out.println("CheckId loi_1: " + previous + " " + current);
				return true; // Danh sách không tăng dần
			}
		}

		return false; // Danh sách tăng dần hoặc không có phần tử
	}

	public static boolean isNotIncreasing(List<Pair<String, Double>> currentList,
			List<Pair<String, Double>> previousList) {
		if (CollectionUtils.isNotEmpty(currentList) && CollectionUtils.isNotEmpty(previousList)) {
			final Pair<String, Double> current = currentList.get(0);
			final Pair<String, Double> previous = previousList.get(previousList.size() - 1);
			final int compareResult = comparePairs(current, previous);
			if (compareResult <= 0) {
				System.out.println("CheckId loi_2: " + previous + " " + current);
				return true; // Danh sách không tăng dần
			}
		}
		return false; // Danh sách tăng dần hoặc không có phần tử
	}

	private static List<String> generateCountSuffixList(String fileName, int count) {
		final List<String> fileList = new ArrayList<>();
		if (StringUtils.isBlank(fileName)) {
			return fileList;
		}
		final String baseName = getBaseName(fileName);
		final String extension = getFileExtension(fileName);

		for (int i = 1; i <= count; i++) {
			final String suffixedName = baseName + "_suff_" + i + extension;
			fileList.add(suffixedName);
		}

		return fileList;
	}

	private static String getBaseName(String fileName) {
		final int dotIndex = fileName.lastIndexOf('.');
		if (dotIndex == -1) {
			return fileName;
		}
		return fileName.substring(0, dotIndex);
	}

	private static String getFileExtension(String fileName) {
		final int dotIndex = fileName.lastIndexOf('.');
		if (dotIndex == -1) {
			return "";
		}
		return fileName.substring(dotIndex);
	}

}

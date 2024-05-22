/**
 *
 */
package com.alxvn.backlog;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Deque;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.alxvn.backlog.behavior.BacklogBehavior;
import com.alxvn.backlog.dto.BacklogDetail;
import com.alxvn.backlog.dto.CustomerTarget;
import com.alxvn.backlog.dto.PjjyujiDetail;
import com.alxvn.backlog.dto.WorkingProcess;
import com.alxvn.backlog.handle.IncorrectFullNameException;
import com.alxvn.backlog.schedule.BacklogExcel;
import com.alxvn.backlog.util.Helper;
import com.alxvn.backlog.util.WorkingReportHelper;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;

/**
 * @author KEDD
 *
 */
@Service
public class BacklogService implements BacklogBehavior {

	private static final Logger log = LoggerFactory.getLogger(BacklogService.class);

	private static final DateTimeFormatter FORMATTER_YYYYMMDD = DateTimeFormatter.ofPattern("yyyy/MM/dd");
	private static final DateTimeFormatter FORMATTER_MMMDDYYYY = DateTimeFormatter.ofPattern("MMM. dd, yyyy",
			Locale.ENGLISH);

	public String stastics(final MultipartFile pjjyujiDataCsv, final MultipartFile backlogIssues)
			throws IOException, CsvException, IncorrectFullNameException {

		List<PjjyujiDetail> pds = new ArrayList<>();
		if (pjjyujiDataCsv != null) {
			final var pjjyujiDataCsvName = pjjyujiDataCsv.getOriginalFilename();
			if (StringUtils.isNotEmpty(pjjyujiDataCsvName) && pjjyujiDataCsvName.endsWith(".csv")) {
				pds = readPjjyujiDataCsv(pjjyujiDataCsv);
			} else {
				//
			}
		}

		List<BacklogDetail> bds = new ArrayList<>();
		if (backlogIssues != null) {
			final var backlogIssuesName = backlogIssues.getOriginalFilename();
			if (StringUtils.isNotEmpty(backlogIssuesName) && backlogIssuesName.endsWith(".csv")) {
				bds = readBacklogIssues(backlogIssues);
			} else {
				//
			}
		}

		return genSchedule(pds, bds);
	}

	public void stastics(final String workingReportFilePath, final String backlogIssuesFilePath)
			throws IOException, CsvException, IncorrectFullNameException {
		log.debug("stastics START");
		final var pds = readPjjyujiDataCsv(workingReportFilePath);
		final var bds = readBacklogIssues(backlogIssuesFilePath);

		genSchedule(pds, bds);
		log.debug("stastics END");
	}

	private final Comparator<? super BacklogDetail> comparator = (o1, o2) -> {
		if (o1 == null && o2 == null) {
			return 0;
		}
		if (o1 == null) {
			return -1;
		}
		if (o2 == null) {
			return 1;
		}
		// So sánh theo trường field1
		if (o1.getKey() == null && o2.getKey() == null) {
			// Nếu cả hai trường field1 đều là null, tiếp tục so sánh trường tiếp theo
		} else if (o1.getKey() == null) {
			return -1;
		} else if (o2.getKey() == null) {
			return 1;
		} else {
			final var result = o1.getKey().compareTo(o2.getKey());
			if (result != 0) {
				return result;
			}
		}

		// Nếu tất cả các trường đều bằng nhau, trả về kết quả là 0
		return 0;
	};

	private String cleanRootFolderSchedule() throws IOException {
		final var folderPath = BacklogExcel.getPathRootFolder(); // Specify the desired folder path

		final var folder = Path.of(folderPath);
		if (Files.exists(folder)) {
			Files.walkFileTree(folder, new SimpleFileVisitor<>() {
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
		} else {
			Files.createDirectories(folder);
		}
		return folderPath;
	}

	@SuppressWarnings("static-access")
	private String genSchedule(final List<PjjyujiDetail> pds, final List<BacklogDetail> bis) throws IOException {
		log.debug("Xử lý tạo file schedule !!!");
		log.debug("Bắt đầu xử lý.");
		final var projectMap = getMapProject(pds, bis);

		final var util = new BacklogExcel();

		if (!util.isUpdateOldSchedule()) {
			cleanRootFolderSchedule();
		}

		for (final Map.Entry<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> entry : projectMap
				.entrySet()) {
			final var key = entry.getKey();
			final var val = entry.getValue();
			/**
			 * Xác định loại khách hàng I-Front / Symphonizer
			 */
			final var projecType = key.getKey();
			final var projectCd = key.getValue();
			final var workingReports = val.getKey();
			final var backlogs = CollectionUtils.emptyIfNull(val.getValue()).stream().sorted(comparator)
					.collect(Collectors.toList());

			/**
			 * Điền thông tin backlog và working report vào file schedule mới
			 */
			if (StringUtils.isBlank(projectCd)) {
				log.debug("Schedule.wr.isNotValid: {}", workingReports);
				log.debug("Schedule.bl.isNotValid: {}", backlogs);
			}
			util.createSchedule(projecType, projectCd, workingReports, backlogs);
		}
		log.debug("Kết thúc xử lý tạo file schedule !!!");
		if (util.isExportZip()) {
			return util.getPathRootFolder();
		}
		return StringUtils.EMPTY;
	}

	public static void zipFolderByPath(final String folderPath, final String outputZipFile) {
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

	public ResponseEntity<InputStreamResource> zipFolder(final String folderPath) {
		log.debug("Bắt đầu xử lý tạo zip file !!!");
		if (StringUtils.isBlank(folderPath)) {
			return ResponseEntity.noContent().build();
		}

		try (var byteArrayOutputStream = new ByteArrayOutputStream()) {
			zipFolder(folderPath, byteArrayOutputStream);
			final var inputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
			final var resource = new InputStreamResource(inputStream);

			final var headers = new HttpHeaders();
			headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
			headers.setContentDisposition(ContentDisposition.builder("attachment").filename("download.zip").build());
			log.debug("Kết thúc xử lý zip file !!!");
			return new ResponseEntity<>(resource, headers, HttpStatus.OK);
		} catch (final IOException e) {
			// Handle exceptions here
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}

	public void zipFolder(final String folderPath, final OutputStream outputStream) throws IOException {
		final var path = Paths.get(folderPath);
		try (var zipOutputStream = new ZipOutputStream(outputStream)) {
			addFolderToZip(zipOutputStream, path, "");
		}
	}

	private void addFolderToZip(final ZipOutputStream zipOutputStream, final Path path, final String basePath)
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

	private Map<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> getMapProject(
			final List<PjjyujiDetail> pds, final List<BacklogDetail> bis) {
//		<PJCD, PJCDJP>,
		final Map<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> result = new HashMap<>();
		for (final BacklogDetail bd : bis) {
			final var pjCdJp = bd.getPjCdJp();
			final var cusTarget = resolveTarget(bd);
			// for test
//			if (!cusTarget.equals(CustomerTarget.NONE) || !StringUtils.equals("05000189", pjCdJp)) {
//				// 05000189
//				// 03010776
//				continue;
//			}
			final Pair<CustomerTarget, String> projectKey = Pair.of(cusTarget, pjCdJp);
			List<PjjyujiDetail> pdList = new ArrayList<>();
			List<BacklogDetail> bdList = new ArrayList<>();
			if (result.containsKey(projectKey)) {
				final var p = result.get(projectKey);
				pdList = p.getLeft();
				bdList = p.getRight();
				bdList.add(bd);
			}
			final var iterator = pds.iterator();
			while (iterator.hasNext()) {
				final var item = iterator.next();
				final var wrProjectCdJp = item.getPjCdJp();
				if (StringUtils.equals(pjCdJp, wrProjectCdJp)) {
					pdList.add(item);
					iterator.remove(); // remove sau khi thỏa điều kiện
				}
			}
			if (!result.containsKey(projectKey)) {
				bdList = new ArrayList<>();
				bdList.add(bd);
			}
			result.put(projectKey, Pair.of(pdList, bdList));
		}

		// Create a TreeMap to store the sorted map
		final Map<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> sortedResult = new TreeMap<>(
				(pair1, pair2) -> {
					// Compare the CustomerTarget objects
					final var targetComparison = pair1.getLeft().compareTo(pair2.getLeft());

					if (targetComparison != 0) {
						return targetComparison;
					}

					// Compare the String objects
					return pair1.getRight().compareTo(pair2.getRight());
				});
		sortedResult.putAll(result);
		return sortedResult;
	}

	private static final String[] workingReportColumns = { //
			"社員番号", //
			"給与社員番号", //
			"社員名", //
			"日付", //
			"プロジェクトコード", //
			"プロジェクトコード（日本）", //
			"プロジェクト名", //
			"固定プロジェクト名", //
			"プロセスコード", //
			"プロセス", //
			"作業内容", //
			"時間内(分)", //
			"普通残業時間(分)", //
			"法定休日残業(分)", //
			"法定祝日残業時間(分)", //
			"深夜残業(分)" //
	};

	private List<PjjyujiDetail> parseWorkingReport(final CSVReader csvReader)
			throws IOException, CsvException, IncorrectFullNameException {
		final List<PjjyujiDetail> result = new ArrayList<>();
		final var header = csvReader.readNext();

		if (header != null) {
			// Create a mapping of column name to column index
			final Map<String, Integer> columnIndexMap = new HashMap<>();
			for (var i = 0; i < header.length; i++) {
				columnIndexMap.put(header[i], i);
			}

			// Find the indices of the desired columns by name
			final var columnIndices = new Integer[workingReportColumns.length];
			for (var i = 0; i < workingReportColumns.length; i++) {
				columnIndices[i] = columnIndexMap.get(workingReportColumns[i]);
			}

			// Process each row
			final var rows = csvReader.readAll();
			for (final String[] row : rows) {
				// Store the values of the desired columns in a map
				final Map<String, String> columnValues = new HashMap<>();
				for (var i = 0; i < columnIndices.length; i++) {
					if (columnIndices[i] != null) {
						columnValues.put(workingReportColumns[i], row[columnIndices[i]]);
					}
				}

				// Access the values by column name
				final var id = columnValues.get("社員番号");
				// line for 給与社員番号
				final var name = columnValues.get("社員名");
				final var date = columnValues.get("日付");
				final var pjCd = columnValues.get("プロジェクトコード");
				final var pjCdJp = columnValues.get("プロジェクトコード（日本）");
				final var pjName = columnValues.get("プロジェクト名");
				// line for 固定プロジェクト名
				final var processCd = columnValues.get("プロセスコード");
				final var processName = columnValues.get("プロセス");
				final var content = columnValues.get("作業内容");
				final var minute1 = columnValues.get("時間内(分)");
				final var minute2 = columnValues.get("普通残業時間(分)");
				final var minute3 = columnValues.get("法定休日残業(分)");
				final var minute4 = columnValues.get("法定祝日残業時間(分)");
				final var minute5 = columnValues.get("深夜残業(分)");

				final var detail = new PjjyujiDetail.Builder()
						/**/
						.setId(id)
						/**/
						.setName(name)
						/**/
						.setMailId(WorkingReportHelper.getMailIdFromFullName(name))
						/**/
						.setTargetYmd(LocalDate.parse(date, FORMATTER_YYYYMMDD))
						/**/
						.setPjCd(pjCd)
						/**/
						.setPjCdJp(pjCdJp)
						/**/
						.setPjName(pjName)
						/**/
						.setProcess(WorkingProcess.of(processCd, processName))
						/**/
						.setContent(content)
						/**/
						.setAnkenNo(Helper.getAnkenNo(content))
						/**/
						.setMinute(WorkingReportHelper.sumMinute(minute1, minute2, minute3, minute4, minute5))
						/**/
						.build();
				result.add(detail);
			}
		}
		return result;
	}

	private List<PjjyujiDetail> readPjjyujiDataCsv(final MultipartFile file)
			throws IOException, CsvException, IncorrectFullNameException {
		try (final var csvReader = new CSVReaderBuilder(
				new BufferedReader(new InputStreamReader(file.getInputStream(), Charset.forName("Shift_JIS"))))
//				.withSkipLines(1) //
				.build();) {
			return parseWorkingReport(csvReader);
		}

	}

	private List<PjjyujiDetail> readPjjyujiDataCsv(final String filePath)
			throws IOException, CsvException, IncorrectFullNameException {
		try (final var csvReader = new CSVReaderBuilder(new BufferedReader(new InputStreamReader(
				BacklogService.class.getClassLoader().getResourceAsStream(filePath), Charset.forName("Shift_JIS"))))
//				.withSkipLines(1) //
				.build();) {
			return parseWorkingReport(csvReader);
		}

	}

	private LocalDate parseBacklogDate(final String str) {
		if (StringUtils.isBlank(str)) {
			return null;
		}
		return LocalDate.parse(str, FORMATTER_MMMDDYYYY);
	}

	private LocalDate parseBacklogCustomDate(final String str) {
		if (StringUtils.isBlank(str)) {
			return null;
		}
		return LocalDate.parse(str, FORMATTER_YYYYMMDD);
	}

	private String extractPjCdFromMileStone(final String ms, final String sj) {
		final var regex = "(\\d{8})";

		final var pattern = Pattern.compile(regex);
		final var matcher = pattern.matcher(ms);

		if (matcher.find()) {
			return matcher.group(1);
		}
		final var matcherSj = pattern.matcher(sj);
		if (matcherSj.find()) {
			return matcherSj.group(1);
		}
		return StringUtils.EMPTY;
	}

	private String extractProcess(final String text) {
		if (StringUtils.isNotBlank(text)) {
			final var regex = "(\\d+)";

			final var pattern = Pattern.compile(regex);
			final var matcher = pattern.matcher(text);

			if (matcher.find()) {
				return matcher.group(1);
			}
		}
		return text;
	}

	private static final String[] backlogColumns = { //
			"ID", //
			"Project ID", //
			"Project Name", //
			"Key ID", //
			"Key", //
			"Issue Type ID", //
			"Issue Type", //
			"Category ID", //
			"Category Name", //
			"Version ID", //
			"Version", //
			"Subject", //
			"Description", //
			"Status ID", //
			"Status", //
			"Priority ID", //
			"Priority", //
			"Milestone ID", //
			"Milestone", //
			"Resolution ID", //
			"Resolution", //
			"Assignee ID", //
			"Assignee", //
			"Create User ID", //
			"Created by", //
			"Created Date", //
			"Parent issue key", //
			"Start Date", //
			"Due date", //
			"Estimated Hours", //
			"Actual Hours", //
			"Update User ID", //
			"Updated by", //
			"Updated", //
			"Attachment", //
			"Shared File", //
			"顧客", // mã dự án
			"開始予定日", //
			"完了予定日", //
			"進捗", //
			"納品予定日", //
			"process(Of Wk Report)", //
			"関連する課題(親)", //
			"課題カテゴリ", //
			"課題発生元", //
			"課題発生者", //
			"課題発生第三者"//
	};

	private BacklogDetail getRecord(final Map<String, String> columnValues) throws IncorrectFullNameException {
		final var status = columnValues.get("Status");
		final var backlogKey = columnValues.get("Key");
		final var issueType = columnValues.get("Issue Type");
		final var category = columnValues.get("Category Name");
		final var subject = columnValues.get("Subject");
		// Priority
		final var milestone = columnValues.get("Milestone");
		final var assignee = columnValues.get("Assignee");
		final var parentKey = columnValues.get("Parent issue key");
		final var actualStartDate = parseBacklogDate(columnValues.get("Start Date"));
		final var actualDueDate = parseBacklogDate(columnValues.get("Due date"));
		final var estimatedHours = columnValues.get("Estimated Hours");
		final var actualHours = columnValues.get("Actual Hours");
		final var targetCustomer = columnValues.get("顧客");
		final var targetTaskId = columnValues.get("課題");
		final var expectedStartDate = parseBacklogCustomDate(columnValues.get("開始予定日"));
		final var expectedDueDate = parseBacklogCustomDate(columnValues.get("完了予定日"));
		final var progress = extractProcess(columnValues.get("進捗"));
		final var expectedDeliveryDate = parseBacklogCustomDate(columnValues.get("納品予定日"));
		final var processOfWr = columnValues.get("process(Of Wk Report)");
		// 関連する課題(親)
		final var bugCategory = columnValues.get("課題カテゴリ");
		final var bugOrigin = columnValues.get("課題発生元");
		final var bugCreator = columnValues.get("課題発生者");
		final var bug3rdTest = columnValues.get("課題発生第三者");

		final var ankenNo = StringUtils.defaultIfBlank(targetTaskId, Helper.getAnkenNo(subject));
		return new BacklogDetail.Builder().key(backlogKey) //
				.issueType(issueType) //
				.ankenNo(ankenNo) //
				.mailId(WorkingReportHelper.getMailIdFromFullName(assignee)) //
				.pjCdJp(extractPjCdFromMileStone(milestone, subject)) //
				.category(category) //
				.subject(subject) //
				.milestone(milestone) //
				.assignee(assignee) //
				.parentKey(parentKey) //
				.expectedStartDate(expectedStartDate) //
				.expectedDueDate(expectedDueDate) //
				.expectedDeliveryDate(expectedDeliveryDate) //
				.estimatedHours(StringUtils.isNotBlank(estimatedHours) ? new BigDecimal(estimatedHours) : null) //
				.actualStartDate(actualStartDate) //
				.actualDueDate(actualDueDate) //
				.actualHours(StringUtils.isNotBlank(actualHours) ? new BigDecimal(actualHours) : null) //
				.status(status) //
				.targetCustomer(targetCustomer) //
				.progress(progress) //
				.bugCategory(bugCategory) //
				.bugOrigin(bugOrigin) //
				.bugCreator(bugCreator) //
				.bug3rdTest(bug3rdTest) //
				.processOfWr(processOfWr)
				/**/
				.build();
	}

	private List<BacklogDetail> parseBacklogDetail(final CSVReader csvReader)
			throws IOException, CsvException, IncorrectFullNameException {
		final List<BacklogDetail> result = new ArrayList<>();
		// Read the header row
		final var header = csvReader.readNext();

		if (header != null) {
			// Create a mapping of column name to column index
			final Map<String, Integer> columnIndexMap = new HashMap<>();
			for (var i = 0; i < header.length; i++) {
				columnIndexMap.put(header[i], i);
			}

			// Find the indices of the desired columns by name
			final var columnIndices = new Integer[backlogColumns.length];
			for (var i = 0; i < backlogColumns.length; i++) {
				columnIndices[i] = columnIndexMap.get(backlogColumns[i]);
			}

			// Process each row
			final var rows = csvReader.readAll();
			for (final String[] row : rows) {
				final Map<String, String> columnValues = new HashMap<>();
				for (final Map.Entry<String, Integer> entry : columnIndexMap.entrySet()) {
					final var key = entry.getKey();
					final var val = entry.getValue();
					columnValues.put(key, row[val]);
				}
				result.add(getRecord(columnValues));
			}
		}
		return result;
	}

	private List<BacklogDetail> readBacklogIssues(final MultipartFile file)
			throws IOException, CsvException, IncorrectFullNameException {
		try (final var csvReader = new CSVReaderBuilder(
				new BufferedReader(new InputStreamReader(file.getInputStream(), StandardCharsets.UTF_8)))
//				.withSkipLines(1) //
				.build();) {
			return parseBacklogDetail(csvReader);
		}
	}

	private List<BacklogDetail> readBacklogIssues(final String filePath)
			throws IOException, CsvException, IncorrectFullNameException {
		try (final var csvReader = new CSVReaderBuilder(new BufferedReader(new InputStreamReader(
				BacklogService.class.getClassLoader().getResourceAsStream(filePath), StandardCharsets.UTF_8)))
//				.withSkipLines(1) //
				.build();) {
			return parseBacklogDetail(csvReader);
		}
	}

	@Override
	public CustomerTarget resolveTarget(final BacklogDetail bd) {
		var cusTarget = CustomerTarget.NONE;
		final var milestone = bd.getMilestone();
		final var targetCustomer = bd.getTargetCustomer();
		if (StringUtils.isBlank(targetCustomer)) {
			if (StringUtils.containsIgnoreCase(milestone, "sym")) {
				cusTarget = CustomerTarget.SYMPHONIZER;
			} else if (StringUtils.containsIgnoreCase(milestone, "i-front")
					|| StringUtils.containsIgnoreCase(milestone, "ifront")) {
				cusTarget = CustomerTarget.IFRONT;
			}
		} else if (StringUtils.containsIgnoreCase(targetCustomer, "sym")) {
			cusTarget = CustomerTarget.SYMPHONIZER;
		} else if (StringUtils.containsIgnoreCase(targetCustomer, "i-front")
				|| StringUtils.containsIgnoreCase(targetCustomer, "ifront")) {
			cusTarget = CustomerTarget.IFRONT;
		} else if (StringUtils.containsIgnoreCase(targetCustomer, "dmp")
				|| StringUtils.containsIgnoreCase(targetCustomer, "katch")) {
			cusTarget = CustomerTarget.DMP;
		}
		return cusTarget;
	}

}

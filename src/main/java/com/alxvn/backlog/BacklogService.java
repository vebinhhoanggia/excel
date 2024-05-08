/**
 *
 */
package com.alxvn.backlog;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.alxvn.backlog.dto.BacklogDetail;
import com.alxvn.backlog.dto.CustomerTarget;
import com.alxvn.backlog.dto.PjjyujiDetail;
import com.alxvn.backlog.dto.WorkingProcess;
import com.alxvn.backlog.handle.IncorrectFullNameException;
import com.alxvn.backlog.util.BacklogExcelUtil;
import com.alxvn.backlog.util.Helper;
import com.alxvn.backlog.util.ScheduleHelper;
import com.alxvn.backlog.util.WorkingReportHelper;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;

/**
 * @author KEDD
 *
 */
@Service
public class BacklogService {

	private static final Logger log = LoggerFactory.getLogger(BacklogService.class);

	private static final DateTimeFormatter FORMATTER_YYYYMMDD = DateTimeFormatter.ofPattern("yyyy/MM/dd");
	private static final DateTimeFormatter FORMATTER_MMMDDYYYY = DateTimeFormatter.ofPattern("MMM. dd, yyyy",
			Locale.ENGLISH);

//	private static final String pathSymTemplate = "\\\\192.168.10.40\\project\\AllexceedJP\\Symphonizer\\Vietnamese\\%s\\05_keikaku";
//	private static final String pathIfrontTemplate = "\\\\192.168.10.40\\project\\AllexceedJP\\i-Front\\Vietnamese\\%s\\05_keikaku";
	private static final String pathSymTemplate = "D:\\Doc\\Backlog\\sym\\%s";
	private static final String pathIfrontTemplate = "D:\\Doc\\Backlog\\ifront\\%s";
	private static final String pathDefaultTemplate = "D:\\Doc\\Backlog\\default\\%s";
	private static final String subFldName = "Backlog_%s";

	public void stastics(final MultipartFile pjjyujiDataCsv, final MultipartFile backlogIssues,
			final MultipartFile backlogGanttChart) throws IOException, CsvException, IncorrectFullNameException {

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

		if (backlogGanttChart != null) {
			final var backlogGanttChartName = backlogGanttChart.getOriginalFilename();
			if (StringUtils.isNotEmpty(backlogGanttChartName)
					&& (backlogGanttChartName.endsWith(".xlsx") || backlogGanttChartName.endsWith(".xls"))) {
				readBacklogGanttChart(backlogGanttChart);
			} else {
				//
			}
		}

		genSchedule(pds, bds);
	}

	public void stastics(final String workingReportFilePath, final String backlogIssuesFilePath)
			throws IOException, CsvException, IncorrectFullNameException {
		log.debug("stastics START");
		final var pds = readPjjyujiDataCsv(workingReportFilePath);
		final var bds = readBacklogIssues(backlogIssuesFilePath);

		genSchedule(pds, bds);
		log.debug("stastics END");
	}

	private void genSchedule(final List<PjjyujiDetail> pds, final List<BacklogDetail> bis) throws IOException {

		final var projectMap = getListProject(pds, bis);
//		bis.stream().filter(x -> StringUtils.containsIgnoreCase(x.getMilestone(), "03010791"))
//				.collect(Collectors.toList());
		for (final Map.Entry<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> entry : projectMap
				.entrySet()) {
			final var key = entry.getKey();
			final var val = entry.getValue();
			/**
			 * Xác định loại khách hàng I-Front / Symphonizer
			 */
			final var projecType = key.getKey();
			final var projectCd = key.getValue();
			String projectScheduleTemplate = null;
			projectScheduleTemplate = switch (projecType) {
			case IFRONT -> pathIfrontTemplate;
			case SYMPHONIZER -> pathSymTemplate;
			default -> pathDefaultTemplate;
			};
			final var workingReports = val.getKey();
			final var backlogs = val.getValue();

			/**
			 * Điền thông tin backlog và working report vào file schedule mới
			 */
			genSchedule(projectScheduleTemplate, projectCd, workingReports, backlogs);
		}
		log.debug("All schedule created successfully.");
	}

	private Map<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> getListProject(
			final List<PjjyujiDetail> pds, final List<BacklogDetail> bis) {
//		<PJCD, PJCDJP>,
		final Map<Pair<CustomerTarget, String>, Pair<List<PjjyujiDetail>, List<BacklogDetail>>> result = new HashMap<>();
		var cusTarget = CustomerTarget.NONE;
		for (final BacklogDetail bd : bis) {
			final var pjCdJp = bd.getPjCdJp();
			if (StringUtils.isBlank(pjCdJp)) {
//				continue;
			}
			final var anken = bd.getAnkenNo();
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
			final Pair<CustomerTarget, String> projectKey = Pair.of(cusTarget, pjCdJp);
			List<PjjyujiDetail> pdList = new ArrayList<>();
			List<BacklogDetail> bdList = new ArrayList<>();
			final var iterator = pds.iterator();
			if (result.containsKey(projectKey)) {
				final var p = result.get(projectKey);
				pdList = p.getLeft();
				bdList = p.getRight();
				bdList.add(bd);
			}
			while (iterator.hasNext()) {
				final var item = iterator.next();
				final var ankenNo = item.getAnkenNo();
				if (StringUtils.equals(anken, ankenNo)) {
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

		// Danh sach cac working report con lai
		for (final PjjyujiDetail pd : pds) {
//			String targetCustomer = null;
//
//			final String pjCdJp = pd.getPjCdJp();
//			final String anken = pd.getAnkenNo();
//			if (StringUtils.contains(anken, "ifront")) {
//				targetCustomer = "I-front";
//			}
//			final Pair<String, String> k = Pair.of(targetCustomer, pjCdJp);
//			List<PjjyujiDetail> pdList = new ArrayList<>();
//			List<BacklogDetail> bdList = new ArrayList<>();
//			final Iterator<BacklogDetail> iterator = bis.iterator();
//			if (result.containsKey(k)) {
//				final Pair<List<PjjyujiDetail>, List<BacklogDetail>> p = result.get(k);
//				pdList = p.getLeft();
//				pdList.add(pd);
//				bdList = p.getRight();
//			}
//			while (iterator.hasNext()) {
//				final BacklogDetail item = iterator.next();
//				final String ankenNo = item.getAnkenNo();
//				if (StringUtils.equals(anken, ankenNo)) {
//					bdList.add(item);
//					iterator.remove(); // Xóa phần tử thỏa mãn điều kiện
//				}
//			}
//			if (!result.containsKey(k)) {
//				pdList = new ArrayList<>();
//				pdList.add(pd);
//			}
//			result.put(k, Pair.of(pdList, bdList));
		}
		return result;
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
			"進捗 Progress", //
			"納品予定日", //
			"process(Of Wk Report)", //
			"関連する課題(親)", //
			"課題カテゴリ", //
			"課題発生元", //
			"課題発生者", //
			"課題発生第三者"//
	};

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
				// Store the values of the desired columns in a map
				final Map<String, String> columnValues = new HashMap<>();
				for (var i = 0; i < columnIndices.length; i++) {
					if (columnIndices[i] != null) {
						columnValues.put(backlogColumns[i], row[columnIndices[i]]);
					}
				}

				// Access the values by column name
				final var status = columnValues.get("Status");
				if (StringUtils.equals(status, "Closed")) {
//					continue;
				}
				final var columnKeyValue = columnValues.get("Key");
				final var issueType = columnValues.get("Issue Type");
				final var category = columnValues.get("Category Name");
				final var subject = columnValues.get("Subject");
//							final String columnPriorityValue = columnValues.get("Priority");
				final var milestone = columnValues.get("Milestone");
				final var assignee = columnValues.get("Assignee");
				final var parentKey = columnValues.get("Parent issue key");
				final var actualStartDate = parseBacklogDate(columnValues.get("Start Date"));
				final var actualDueDate = parseBacklogDate(columnValues.get("Due date"));
				final var estimatedHours = columnValues.get("Estimated Hours");
				final var actualHours = columnValues.get("Actual Hours");
				final var targetCustomer = columnValues.get("顧客");
				final var expectedStartDate = parseBacklogCustomDate(columnValues.get("開始予定日"));
				final var expectedDueDate = parseBacklogCustomDate(columnValues.get("完了予定日"));
				final var progress = columnValues.get("進捗 Progress");
				final var expectedDeliveryDate = parseBacklogCustomDate(columnValues.get("納品予定日"));
				final var processOfWr = columnValues.get("process(Of Wk Report)");
//							final String column関連する課題Value = columnValues.get("関連する課題(親)");
				final var bugCategory = columnValues.get("課題カテゴリ");
				final var bugOrigin = columnValues.get("課題発生元");
				final var bugCreator = columnValues.get("課題発生者");
				final var bug3rdTest = columnValues.get("課題発生第三者");

				final var detail = new BacklogDetail.Builder().key(columnKeyValue) //
						.issueType(issueType) //
						.ankenNo(Helper.getAnkenNo(subject)) //
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
				result.add(detail);
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

	private List<String> readBacklogGanttChart(final MultipartFile file) {
		final List<String> fileData = new ArrayList<>();

//		final Workbook workbook = WorkbookFactory.create(file.getInputStream());
//		final Sheet sheet = workbook.getSheetAt(0);
//		for (final Row row : sheet) {
//			for (final Cell cell : row) {
//				final String cellValue = cell.toString();
//				fileData.add(cellValue);
//			}
//		}
//		workbook.close();

		return fileData;
	}

	/**
	 * Tạo thư mục chứa schedule được tạo từ backlog
	 *
	 * @param schedulePath
	 * @return
	 * @throws IOException
	 */
	private Path createFolderBacklogSchedule(final Path schedulePath) throws IOException {
		final var formatter = DateTimeFormatter.ofPattern("yyyyMM");

		// Format the LocalDateTime object to a string using the formatter
		final var folderName = LocalDateTime.now().format(formatter);
		final var subfolderPath = schedulePath.resolve(String.format(subFldName, folderName));
		// Create the subfolder
		Files.createDirectories(subfolderPath);
		return subfolderPath;
	}

	private Path backupSch(final Path schedulePath) throws IOException {
		final var subfolderPath = schedulePath.resolve(String.format(subFldName, LocalDateTime.now()));
		// Create the subfolder
		Files.createDirectory(subfolderPath);

		try (var walkStream = Files.walk(schedulePath)) {
			walkStream.filter(Files::isRegularFile) // Only take files
					.filter(file -> ScheduleHelper.isValidSchFileName(file.toFile())) //
					.forEach(file -> {
						try {
							Files.copy(file, subfolderPath.resolve(file.getFileName())); //
						} catch (final IOException e) {
							throw new IllegalArgumentException("Backup schedule:", e);
						}
					});
		}
		return subfolderPath;
	}

	private void genSchedule(final String schPathTemplate, final String pjcd, final List<PjjyujiDetail> pds,
			final List<BacklogDetail> bds) throws IOException {
		if (StringUtils.isNotBlank(pjcd)) {
			final var schedulePath = Paths.get(String.format(schPathTemplate, pjcd));
			final var backlogSchedulePath = createFolderBacklogSchedule(schedulePath);
			final var util = new BacklogExcelUtil();
			util.createSchedule(pjcd, backlogSchedulePath, pds, bds);
		} else {
			log.debug("Schedule.wr.isNotValid: {}", pds);
			log.debug("Schedule.bl.isNotValid: {}", bds);
			final var schedulePath = Paths.get(String.format(schPathTemplate, pjcd));
			final var backlogSchedulePath = createFolderBacklogSchedule(schedulePath);
			final var util = new BacklogExcelUtil();
			util.createSchedule(pjcd, backlogSchedulePath, pds, bds);
		}
	}

	public void updateSchedule(final String schPathTemplate, final String pjcd,
			final Pair<List<PjjyujiDetail>, List<BacklogDetail>> data) throws IOException {
		final var schedulePath = Paths.get(String.format(schPathTemplate, pjcd));
		// Check if the file or directory exists
		final var exists = Files.exists(schedulePath);
		if (!exists) {
			return;
		}
		final var targetPath = backupSch(schedulePath);
		final var pds = data.getLeft();
		final var bds = data.getRight();

		final List<PjjyujiDetail> pdSame = new ArrayList<>();
		final List<BacklogDetail> bdSame = new ArrayList<>();

		final var pdt = pds.iterator();
		while (pdt.hasNext()) {
			final var item = pdt.next();
			final var ankenNo = item.getAnkenNo();
			if (bds.stream().anyMatch(x -> StringUtils.equals(ankenNo, x.getAnkenNo()))) {
				pdSame.add(item);
				pdt.remove(); // Xóa phần tử thỏa mãn điều kiện
			}
		}

		final var bdt = bds.iterator();
		while (bdt.hasNext()) {
			final var item = bdt.next();
			final var ankenNo = item.getAnkenNo();
			if (pdSame.stream().anyMatch(x -> StringUtils.equals(ankenNo, x.getAnkenNo()))) {
				bdSame.add(item);
				bdt.remove(); // Xóa phần tử thỏa mãn điều kiện
			}
		}
		// TODO: xu ly du lieu dau vao
		// List PjjyujiDetail, BacklogDetail ton tai o backlog, wr => pdSame, bdSame
		// List PjjyujiDetail khong ton tai o backlog => pds
		// List BacklogDetail khong ton tai o wr => bds
		try (var fileStream = Files.walk(targetPath)) {
			fileStream.filter(Files::isRegularFile).forEach(filePath -> {
				try (InputStream inputStream = new BufferedInputStream(Files.newInputStream(filePath));
						var workbook = WorkbookFactory.create(inputStream)) {
					if (Objects.isNull(workbook)) {
						log.debug("Schedule.isNotValid: {}", filePath);
						return;
					}
					final var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
					final var columnToSearch = 2; // Column index (0-based) to search
					final var searchText = "";
					for (final Sheet sheet : workbook) {
						for (final CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
							if (mergedRegion.isInRange(0, columnToSearch)) {
								for (var row = mergedRegion.getFirstRow(); row <= mergedRegion.getLastRow(); row++) {
									final var sheetRow = sheet.getRow(row);
									if (sheetRow != null) {
										final var cell = sheetRow.getCell(columnToSearch);
										if (cell != null && cell.getCellType() == CellType.STRING) {
											final var cellValue = cell.getStringCellValue();
											if (cellValue.equals(searchText)) {
												System.out.println("Merged cell found: " + mergedRegion);
												break;
											}
										}
									}
								}
							}
						}

					}

				} catch (EncryptedDocumentException | IOException e) {
					throw new IllegalArgumentException("update schedule:", e);
				}
			});
		}
	}

}

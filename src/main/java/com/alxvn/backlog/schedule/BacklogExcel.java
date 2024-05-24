/**
 *
 */
package com.alxvn.backlog.schedule;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Predicate;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alxvn.backlog.BacklogService;
import com.alxvn.backlog.behavior.GenSchedule;
import com.alxvn.backlog.dto.BacklogDetail;
import com.alxvn.backlog.dto.CustomerTarget;
import com.alxvn.backlog.dto.PjjyujiDetail;
import com.alxvn.backlog.dto.WorkingPhase;
import com.alxvn.backlog.util.ScheduleHelper;

/**
 *
 */
public class BacklogExcel implements GenSchedule {

	private static final Logger log = LoggerFactory.getLogger(BacklogExcel.class);

	private static final String COLUMN_A_CHARACTER = "A";
	private static final String COLUMN_B_CHARACTER = "B";
	private static final String COLUMN_ANKEN_CHARACTER = "C";
	private static final String COLUMN_SCREEN_CHARACTER = "D";
	private static final String COL_OPERATION_CHAR = "E";
	private static final String COLUMN_PIC_CHARACTER = "F";
	private static final String COLUMN_TOTAL_CHARACTER = "F";
	private static final String COLUMN_STATUS_CHARACTER = "G";
	private static final String COL_EXPECT_HOURS_CHAR = "H";
	private static final String COL_EXPECT_START_YMD_CHAR = "I";
	private static final String COL_EXPECT_END_YMD_CHAR = "J";
	private static final String COL_EXPECT_DELIVERY_YMD_CHAR = "K";
	private static final String COL_BACKLOG_ID_CHAR = "L";
	private static final String COL_ACTUAL_TOTAL_HOURS_BACKLOG_CHAR = "M";
	private static final String COL_ACT_HOURS_CHAR = "N";
	private static final String COL_ACT_START_YMD_CHAR = "O";
	private static final String COL_ACT_END_YMD_CHAR = "P";
	private static final String COL_ACT_PROGRESS_CHAR = "Q";
	private static final String COL_ACT_DELIVERY_YMD_CHAR = "R";
	private static final String COL_TEMPLATE_START_DATE = "S";
	private static final int ROW_PJ_NO_IDX = 3;
	private static final int ROW_PJ_CD_IDX = 4;
	private static final int COL_PJ_NO_IDX = 5;

	private static final String TOTAL_CHARACTER = "Total";
	private static final int COLUMN_A_INDEX = CellReference.convertColStringToIndex(COLUMN_A_CHARACTER);
	private static final int COLUMN_B_INDEX = CellReference.convertColStringToIndex(COLUMN_B_CHARACTER);
	private static final int COLUMN_ANKEN_INDEX = CellReference.convertColStringToIndex(COLUMN_ANKEN_CHARACTER);
	private static final int COLUMN_SCREEN_INDEX = CellReference.convertColStringToIndex(COLUMN_SCREEN_CHARACTER);
	private static final int COLUMN_STATUS_INDEX = CellReference.convertColStringToIndex(COLUMN_STATUS_CHARACTER);
	private static final int TARGET_MONTH_ROW_IDX = 7;
	private static final int TARGET_DATE_ROW_IDX = 8;
	private static final int COLUMN_START_DATE_INPUT_IDX = CellReference
			.convertColStringToIndex(COL_TEMPLATE_START_DATE);
	private static final DateTimeFormatter FORMATTER_YYYYMMDD = DateTimeFormatter.ofPattern("yyyy/MM/dd");
	private static final DateTimeFormatter FORMATTER_YYYYMMDDHHMMSS = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

	private static final String TEMPLATE_FILE = "QDA-0222a_プロジェクト管理表_{projectCd}_{range}.xlsm";
	private static final String TEMPLATE_TOTAL_ACT_HOURS = "SUM({cNameS}{rIdx}:{cNameE}{rIdx})";
	private static final String TEMPLATE_NEXT_DATE_FORMULA = "{preCol}+1";

	private static final String SCHEDULE_TEMPLATE_PATH = "templates/QDA-0222a_プロジェクト管理表.xlsm";

	private static final String PATH_ROOT_FOLDER = "D:\\Doc\\Backlog";
	private static final String PATH_SYM_TEMPLATE = "%s\\Target_%s\\sym\\%s";
	private static final String PATH_IFRONT_TEMPLATE = "%s\\Target_%s\\ifront\\%s";
	private static final String PATH_DMP_TEMPLATE = "%s\\Target_%s\\dmp\\%s";
	private static final String PATH_DEFAULT_TEMPLATE = "%s\\Target_%s\\default\\%s";

	private final String executeTime;
	private String targetFolder = null;

	public BacklogExcel(final LocalDateTime now) {
		executeTime = now.format(FORMATTER_YYYYMMDDHHMMSS);
		targetFolder = String.format("%s\\Target_%s", PATH_ROOT_FOLDER, executeTime);
	}

	public String getPathRootFolder() {
		return targetFolder;
	}

	public String getSymTemplatePath(final String projectCd) {
		return String.format(PATH_SYM_TEMPLATE, PATH_ROOT_FOLDER, executeTime, projectCd);
	}

	public String getiFrontTemplatePath(final String projectCd) {
		return String.format(PATH_IFRONT_TEMPLATE, PATH_ROOT_FOLDER, executeTime, projectCd);
	}

	public String getDmpTemplatePath(final String projectCd) {
		return String.format(PATH_DMP_TEMPLATE, PATH_ROOT_FOLDER, executeTime, projectCd);
	}

	public String getDefaultTemplatePath(final String projectCd) {
		return String.format(PATH_DEFAULT_TEMPLATE, PATH_ROOT_FOLDER, executeTime, projectCd);
	}

	private boolean compareCellRangeAddresses(final CellRangeAddress range1, final CellRangeAddress range2) {
		// Compare the first row, last row, first column, and last column
		return range1.getFirstRow() == range2.getFirstRow() && range1.getLastRow() == range2.getLastRow();
	}

	private boolean isTotalRow(final Row row) {
		final var columnTotalIndex = CellReference.convertColStringToIndex(COLUMN_TOTAL_CHARACTER);
		return ScheduleHelper.isTotalRow(row, TOTAL_CHARACTER, columnTotalIndex);
	}

	/*
	 * Cập nhật công thức tính tổng dựa trên việc thêm cột mới.
	 */
	private void updatedTotalActualHoursFormula(final Sheet sheet) {
		final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var columnIndex = CellReference.convertColStringToIndex(COL_ACT_HOURS_CHAR);
		for (final Row row : sheet) {
			if (isTotalRow(row)) {
				break;
			}
			final var chr = ScheduleHelper.convertColumnIndexToName(row.getLastCellNum());
			final var rNum = row.getRowNum();
			if (rNum >= 9) {
				final var cell = row.getCell(columnIndex);
				if (cell != null) {
					final var adjustedFormula = StringUtils.replaceEach(TEMPLATE_TOTAL_ACT_HOURS,
							new String[] { "{rIdx}", "{cNameS}", "{cNameE}" },
							new String[] { String.valueOf(rNum + 1), COL_TEMPLATE_START_DATE, chr });
					cell.setCellFormula(adjustedFormula);
					formulaEvaluator.evaluate(cell);
				}
			}
		}
	}

	private void setDefaultValForRow(final Row curRow) {
		var curCel = getCell(curRow, COLUMN_STATUS_CHARACTER);
		curCel.setCellValue("未着手");
		curCel = getCell(curRow, COL_EXPECT_HOURS_CHAR);
		curCel.setCellValue(0);
		curCel = getCell(curRow, COL_ACTUAL_TOTAL_HOURS_BACKLOG_CHAR);
		curCel.setCellValue(0);
		curCel = getCell(curRow, COL_ACT_PROGRESS_CHAR);
		curCel.setCellValue(0);
	}

	/*
	 *
	 */
	private void cloneRowFormat(final Row sourceRow, final Row newRow) {
		// Iterate over the cells in the source row
		for (int column = sourceRow.getFirstCellNum(); column <= sourceRow.getLastCellNum(); column++) {
			final var sourceCell = sourceRow.getCell(column); // Get the source cell
			final var newCell = newRow.createCell(column); // Create a new cell in the new row
			if (sourceCell != null) {
				final var sourceCellStyle = sourceCell.getCellStyle(); // Get the cell style of the source cell
				newCell.setCellStyle(sourceCellStyle); // Set the cell style to the new cell
				final var newCellStyle = newCell.getCellStyle(); // Get the cell style of the source cell
				if (column == CellReference.convertColStringToIndex(COL_ACT_HOURS_CHAR)) {
					newCellStyle.setBorderTop(BorderStyle.THIN);
					newCellStyle.setBorderBottom(BorderStyle.THIN);
				}
			}
		}
		setDefaultValForRow(newRow);
	}

	private void setValForMergeCell(final Sheet sheet, final CellRangeAddress mergeCellRange, final int colIdx,
			final String strVal) {
		final var firstRow = mergeCellRange.getFirstRow();
		final var fRow = firstRow;
		// Create a new cell within the merged region
		var r = sheet.getRow(fRow);
		if (r == null) {
			r = sheet.createRow(fRow);
		}
		var c = r.getCell(colIdx);
		if (c == null) {
			c = r.createCell(colIdx);
		}
		// Set the value for the cell
		c.setCellValue(strVal);
	}

	private boolean isHeader(final Row row) {
		return row.getRowNum() <= TARGET_DATE_ROW_IDX;
	}

	private Date toDate(final LocalDate localDate) {
		if (localDate == null) {
			return null;
		}
		// Convert LocalDate to Date
		final var localDateTime = localDate.atStartOfDay();
		final var zonedDateTime = localDateTime.atZone(ZoneId.systemDefault());
		return Date.from(zonedDateTime.toInstant());
	}

	private Optional<WorkingPhase> getOperation(final BacklogDetail backlogDetail) {
		final var processOfWr = backlogDetail.getProcessOfWr();
		final var processOfWrCd = extracProcessOfWrCd(processOfWr);
		if (StringUtils.isBlank(processOfWrCd)) {
			return Optional.ofNullable(null);
		}
		return Optional.ofNullable(WorkingPhase.fromString(processOfWrCd));
	}

	private double getProgress(final String backlogProgress) {
		if (StringUtils.isBlank(backlogProgress)) {
			return 0;
		}
		final double value = NumberUtils.toInt(backlogProgress);
		final var percent = 0.01;

		return value * percent;
	}

	private Cell getCell(final Row curRow, final String colChar) {
		final var cellIdx = CellReference.convertColStringToIndex(colChar);
		var curCel = curRow.getCell(cellIdx);
		if (curCel == null) {
			curCel = curRow.createCell(cellIdx);
		}
		return curCel;
	}

	private void fillBacklogDataForRow(final Row curRow, final BacklogDetail backlogDetail,
			final AtomicInteger indexNo) {
		// "No"
		var curCel = getCell(curRow, COLUMN_A_CHARACTER);
		curCel.setCellValue(indexNo.get());
		// "グループ Group"
		final var parentKey = Optional.ofNullable(backlogDetail).map(BacklogDetail::getParentKey)
				.filter(StringUtils::isNotBlank).orElse(Optional.ofNullable(backlogDetail).map(BacklogDetail::getKey)
						.filter(StringUtils::isNotBlank).orElse(StringUtils.EMPTY));
		if (StringUtils.isNotBlank(parentKey)) {
			curCel = getCell(curRow, COLUMN_B_CHARACTER);
			curCel.setCellValue(parentKey);
		}
		// "画面ID Screen ID"
		final var ankenNo = Optional.ofNullable(backlogDetail).map(BacklogDetail::getAnkenNo).orElse(StringUtils.EMPTY);
		if (StringUtils.isNotBlank(ankenNo)) {
			curCel = getCell(curRow, COLUMN_ANKEN_CHARACTER);
			curCel.setCellValue(ankenNo);
		}
		// 工程 Operation
		final var operation = getOperation(backlogDetail).map(WorkingPhase::getName).orElse(StringUtils.EMPTY);
		curCel = getCell(curRow, COL_OPERATION_CHAR);
		curCel.setCellValue(operation);
		// 担当 PIC
		final var pic = backlogDetail.getMailId();
		curCel = getCell(curRow, COLUMN_PIC_CHARACTER);
		curCel.setCellValue(pic);

		// "予定 Schedule"
		// 工数 Hours
		curCel = getCell(curRow, COL_EXPECT_HOURS_CHAR);
		curCel.setCellValue(
				Optional.ofNullable(backlogDetail.getEstimatedHours()).orElse(BigDecimal.ZERO).doubleValue());
		// 開始日 Begin
		curCel = getCell(curRow, COL_EXPECT_START_YMD_CHAR);
		curCel.setCellValue(toDate(backlogDetail.getExpectedStartDate()));
		// 完了日 End
		curCel = getCell(curRow, COL_EXPECT_END_YMD_CHAR);
		curCel.setCellValue(toDate(backlogDetail.getExpectedDueDate()));
		// 納品日 Delivery
		curCel = getCell(curRow, COL_EXPECT_DELIVERY_YMD_CHAR);
		curCel.setCellValue(toDate(backlogDetail.getExpectedDeliveryDate()));

		// Backlog Information
		// Key
		curCel = getCell(curRow, COL_BACKLOG_ID_CHAR);
		curCel.setCellValue(backlogDetail.getKey());
		// Hours
		curCel = getCell(curRow, COL_ACTUAL_TOTAL_HOURS_BACKLOG_CHAR);
		curCel.setCellValue(Optional.ofNullable(backlogDetail.getActualHours()).orElse(BigDecimal.ZERO).doubleValue());

		// "実績 Actual"
		// 開始日 Begin
		curCel = getCell(curRow, COL_ACT_START_YMD_CHAR);
		curCel.setCellValue(toDate(backlogDetail.getActualStartDate()));
		// 完了日 End
		curCel = getCell(curRow, COL_ACT_END_YMD_CHAR);
		curCel.setCellValue(toDate(backlogDetail.getActualDueDate()));
		// 進捗 Progress
		curCel = getCell(curRow, COL_ACT_PROGRESS_CHAR);
		curCel.setCellValue(getProgress(backlogDetail.getProgress()));
		// 納品日 Delivery
		curCel = getCell(curRow, COL_ACT_DELIVERY_YMD_CHAR);
		curCel.setCellValue(toDate(backlogDetail.getActualDeliveryDate()));
	}

	private Collection<PjjyujiDetail> fillDataForRow(final Sheet sheet, final int curRowIdx,
			final BacklogDetail backlogDetail, final List<PjjyujiDetail> wrTargets, final AtomicInteger indexNo) {
		if (backlogDetail == null) {
			return Collections.emptyList();
		}
		final var curRow = sheet.getRow(curRowIdx);

		fillBacklogDataForRow(curRow, backlogDetail, indexNo);

		final var ankenNo = Optional.ofNullable(backlogDetail).map(BacklogDetail::getAnkenNo).orElse(StringUtils.EMPTY);
		final var pic = backlogDetail.getMailId();
		final var operation = getOperation(backlogDetail).map(WorkingPhase::getName).orElse(StringUtils.EMPTY);

		// fill working report data
		return fillWrData(sheet, curRow, ankenNo, pic, operation, wrTargets);
	}

	private Collection<PjjyujiDetail> fillWrData(final Sheet sheet, final Row curRow, final String ankenNo,
			final String pic, final String operation, final List<PjjyujiDetail> wrTargets) {
		final var wrOfRow = CollectionUtils.emptyIfNull(wrTargets).stream().filter(w -> {
			final var ticketNo = w.getAnkenNo();
			final var mailId = w.getMailId();
			final var processCd = w.getProcess().getCode();
			final var wrOpeationName = Optional.ofNullable(WorkingPhase.fromString(processCd))
					.map(WorkingPhase::getName).orElse(StringUtils.EMPTY);
			return StringUtils.equals(ankenNo, ticketNo) && StringUtils.equals(pic, mailId)
					&& StringUtils.equals(wrOpeationName, operation);
		}).toList();

		final Map<LocalDate, Integer> groupedData = wrOfRow.stream().collect(
				Collectors.groupingBy(PjjyujiDetail::getTargetYmd, Collectors.summingInt(PjjyujiDetail::getMinute)));
		if (MapUtils.isNotEmpty(groupedData)) {
			final var targetDateRow = sheet.getRow(TARGET_DATE_ROW_IDX);
			final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
			for (final Cell c : curRow) {
				final var curColIdx = c.getColumnIndex();
				if (curColIdx < COLUMN_START_DATE_INPUT_IDX) {
					continue;
				}
				final var cellTargetDate = targetDateRow.getCell(curColIdx);
				final var cellTargetVal = ScheduleHelper.getCellValueAsString(cellTargetDate,
						formulaEvaluator.evaluate(cellTargetDate));
				if (StringUtils.isNotBlank(cellTargetVal)) {
					final var targetYmd = LocalDate.parse(cellTargetVal, FORMATTER_YYYYMMDD);
					final var minutes = groupedData.getOrDefault(targetYmd, null);
					if (minutes != null) {
						final var hours = minutes / 60.0; // Convert minutes to hours
						c.setCellValue(hours);
					}
				}
			}
		}
		return CollectionUtils.emptyIfNull(wrOfRow);
	}

	private void standardizedRangeInput(final Sheet sheet, final YearMonth targetYmS, final YearMonth targetYmE) {
		final var row = sheet.getRow(TARGET_MONTH_ROW_IDX);
		String lastTarget = null;
		for (final Cell cell : row) {
			if (cell == null || cell.getColumnIndex() < COLUMN_START_DATE_INPUT_IDX) {
				continue;
			}
			final var cellVal = StringUtils.trim(ScheduleHelper.readContentCell(sheet, cell));
			if (StringUtils.isNotBlank(cellVal)) {
				lastTarget = cellVal;
			}
		}
		if (StringUtils.isBlank(lastTarget)) { // Check sheet is from template
			var currentYm = targetYmS;
			Boolean isFirst = true;
			while (currentYm.isBefore(targetYmE.plusMonths(1))) {
				addColInput(sheet, currentYm, isFirst);
				currentYm = currentYm.plusMonths(1);
				isFirst = false;
			}
		} else {
			clearOldValueWr(sheet, targetYmS, targetYmE);
			// Sheet is old schedule
			// Parse the date string to LocalDate
			final var date = LocalDate.parse(lastTarget, DateTimeFormatter.ofPattern("yyyy/MM/dd"));

			// Extract the YearMonth from LocalDate
			final var yearMonth = YearMonth.from(date).plusMonths(1);
			var currentYm = yearMonth;
			while (currentYm.isBefore(targetYmE.plusMonths(1))) {
				addColInput(sheet, currentYm, false);
				currentYm = currentYm.plusMonths(1);
			}
		}
	}

	private void clearOldValueWr(final Sheet sheet, final YearMonth targetYmS, final YearMonth targetYmE) {
		final var row = sheet.getRow(TARGET_MONTH_ROW_IDX);
		final var formulaEvaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var colStartEnd = ScheduleHelper.findRangeColIndex(formulaEvaluator, sheet, targetYmS, targetYmE);
		if (colStartEnd == null) {
			return;
		}
		final int colStartIdx = colStartEnd.getLeft();
		final int colEndIdx = colStartEnd.getRight();
		for (final Row r : sheet) {
			if (isHeader(r)) {
				continue;
			}
			for (final Cell cell : r) {
				if (cell == null) {
					continue;
				}

				final var curColIdx = cell.getColumnIndex();
				if (curColIdx < colStartIdx) {
					continue;
				}
				if (curColIdx <= colEndIdx) {
					cell.setCellValue((String) null);
				}
				if (curColIdx > colEndIdx) {
					break;
				}
			}
			if (isTotalRow(r)) {
				break;
			}

		}

	}

	private void addColInput(final Sheet sheet, final YearMonth targetYm, final boolean isFirstTargetMonth) {
		final var row = sheet.getRow(TARGET_MONTH_ROW_IDX);
		final var lastColumnIndex = row.getLastCellNum() - 1;

		// Check if there are any rows in the sheet
		if (sheet.getLastRowNum() < 0) {
			log.debug("Sheet {} is empty", sheet.getSheetName());
			return;
		}
		final var lengthOfMonth = targetYm.lengthOfMonth();

		// Loop through each row and copy cell value and style
		final var desiredDay = 1;
		final var localDate = targetYm.atDay(desiredDay);
		for (var i = TARGET_MONTH_ROW_IDX; i <= sheet.getLastRowNum(); i++) {
			final var sourceRow = sheet.getRow(i);
			var destinationRow = sheet.getRow(i);
			if (sourceRow != null) {
				final var sourceCell = sourceRow.getCell(lastColumnIndex);
				if (sourceCell != null && i <= TARGET_DATE_ROW_IDX && isFirstTargetMonth) { // first target month
					sourceCell.setCellValue(localDate);
				}

				Cell destinationCell = null;

				// Check if destination row exists, create it if not
				if (destinationRow == null) {
					destinationRow = sheet.createRow(i);
				}
				var preColStr = CellReference.convertNumToColString(lastColumnIndex);
				// Create new cell in destination column
				final var newColCnt = lengthOfMonth - (isFirstTargetMonth ? 1 : 0);
				for (var j = 1; j <= newColCnt; j++) {
					final var colIdx = lastColumnIndex + j;
					destinationCell = destinationRow.createCell(colIdx);
					// fill target month
					if (!isFirstTargetMonth && j == 1 && i <= TARGET_DATE_ROW_IDX) {
						destinationCell.setCellValue(localDate);
					}
					// fill formula plus date
					if (i == TARGET_DATE_ROW_IDX) {
						final var adjustedFormula = StringUtils.replaceEach(TEMPLATE_NEXT_DATE_FORMULA,
								new String[] { "{preCol}", }, new String[] { preColStr + (TARGET_DATE_ROW_IDX + 1) });
						destinationCell.setCellFormula(adjustedFormula);
					}

					final var newWidth = 4; // Desired width in characters
					sheet.setColumnWidth(colIdx, newWidth * 256); // POI uses units of 1/256th of a character

					// Copy cell style
					if (sourceCell != null) {
						final var newStyle = sheet.getWorkbook().createCellStyle();
						newStyle.setAlignment(HorizontalAlignment.CENTER);
						newStyle.cloneStyleFrom(sourceCell.getCellStyle());
						destinationCell.setCellStyle(newStyle);
					}
					preColStr = CellReference.convertNumToColString(colIdx);
				}
				// merge cell for target month
				final var colIdxS = lastColumnIndex + (isFirstTargetMonth ? 0 : 1);
				final var colIdxE = colIdxS + lengthOfMonth - 1;
				if (i == TARGET_MONTH_ROW_IDX) {
					final var newMergedRegion = new CellRangeAddress(i, i, colIdxS, colIdxE);
					sheet.addMergedRegion(newMergedRegion);
				}

				// set value for merge cell target month

				// set value for new date
			}
		}
	}

	/**
	 * @param args
	 */
	public static void main(final String[] args) {
		final var obj = new BacklogExcel(LocalDateTime.now());
		final var wrPath = "templates/pjjyuji_data_csv_20240509.csv";
		final var backlogPath = "templates/Backlog-Issues-20240514-1217.csv";
		obj.createScheduleFromBacklog(wrPath, backlogPath);
	}

	public void createScheduleFromBacklog(final String wrPath, final String backlogPath) {
		final var backlogService = new BacklogService();
		try {
			backlogService.stastics(wrPath, backlogPath);
		} catch (final Exception e) {
			e.printStackTrace();
		}
	}

	private String getBacklogKeyVal(final Row row) {
		final var formulaEvaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var dataFormatter = new DataFormatter();

		return StringUtils.trim(StringUtils.defaultString(dataFormatter.formatCellValue(
				row.getCell(CellReference.convertColStringToIndex(COL_BACKLOG_ID_CHAR)), formulaEvaluator)));
	}

	private String getBacklogParentKeyVal(final Row row) {
		final var formulaEvaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var dataFormatter = new DataFormatter();

		return StringUtils.trim(StringUtils.defaultString(dataFormatter.formatCellValue(
				row.getCell(CellReference.convertColStringToIndex(COLUMN_B_CHARACTER)), formulaEvaluator)));
	}

	private void shiftRow(final Sheet sheet, final int startRowShift, final int numberOfRowsToShift) {
		if (startRowShift <= sheet.getLastRowNum()) {
			sheet.shiftRows(startRowShift, sheet.getLastRowNum(), numberOfRowsToShift);
		} else {
			log.debug("Sheet {} startRowShift {} numberOfRowsToShift {}", sheet.getSheetName(), startRowShift,
					numberOfRowsToShift);
		}
	}

	private final Predicate<BacklogDetail> isBacklogDetail = b -> Objects.nonNull(b)
			&& !StringUtils.equals(b.getParentKey(), b.getKey());

	private final Predicate<BacklogDetail> isBacklogParent = b -> Objects.nonNull(b)
			&& StringUtils.isBlank(b.getParentKey());

	private List<BacklogDetail> getTargetBacklogs(final Map<String, List<BacklogDetail>> groupedBacklogs,
			final String parentKey) {
		return CollectionUtils.emptyIfNull(groupedBacklogs.get(parentKey)).stream().filter(isBacklogDetail).toList();
	}

	private Cell getCol(final Sheet sheet, final int rowIdx, final int colIdx) {
		var row = sheet.getRow(rowIdx);
		if (row == null) {
			row = sheet.createRow(rowIdx);
		}
		var col = row.getCell(colIdx);
		if (col == null) {
			col = row.createCell(colIdx);
		}
		return col;
	}

	private void fillCommonInfo(final Sheet sheet, final Map<String, List<BacklogDetail>> groupedBacklogs,
			final List<PjjyujiDetail> allWrDatas) {
		groupedBacklogs.entrySet().stream().findFirst()
				.ifPresent(x -> CollectionUtils.emptyIfNull(x.getValue()).stream().findFirst().ifPresent(k -> {
					CollectionUtils.emptyIfNull(allWrDatas).stream().findFirst().map(PjjyujiDetail::getPjCd)
							.filter(StringUtils::isNotBlank).ifPresent(val -> {
								// PJCODE :
								final var col = getCol(sheet, ROW_PJ_NO_IDX, COL_PJ_NO_IDX);
								col.setCellValue(val);
							});
					Optional.of(k).map(BacklogDetail::getPjCdJp).filter(StringUtils::isNotBlank).ifPresent(val -> {
						// PJ-NO ：
						final var col = getCol(sheet, ROW_PJ_CD_IDX, COL_PJ_NO_IDX);
						col.setCellValue(val);
					});
				}));
	}

	private int fillWhenExistsSingleRecord(final Sheet sheet, final Row row, final int curIdx,
			final AtomicInteger indexNo, final List<BacklogDetail> backlogs, final List<PjjyujiDetail> allWrDatas) {

		final var curRowCnt = 1;
		final var curBacklogKeyVal = getBacklogKeyVal(row);
		final var curBacklogParentKeyVal = getBacklogParentKeyVal(row);
		final var newBacklogs = backlogs.stream().filter(x -> !StringUtils.equals(x.getKey(), curBacklogKeyVal))
				.toList();
		final var curBacklog = backlogs.stream().filter(x -> StringUtils.equals(x.getKey(), curBacklogKeyVal))
				.findFirst().orElse(null);

		final var isOnlyParentRow = curBacklog == null && StringUtils.isNotBlank(curBacklogParentKeyVal);
		var parentRowCnt = 0;
		if (isOnlyParentRow) {
			parentRowCnt = 1;
		}
		// Dịch chuyển các dòng
		final var startRowShift = curIdx + curRowCnt;
		final var numberOfRowsToShift = newBacklogs.size() - parentRowCnt;
		if (numberOfRowsToShift > 0) {
			shiftRow(sheet, startRowShift, numberOfRowsToShift);
			// Tạo dòng mới sau khi dịch chuyển
			for (var j = startRowShift; j < startRowShift + numberOfRowsToShift; j++) {
				final var newRow = sheet.createRow(j);
				cloneRowFormat(row, newRow);
			}
		}

		// Cập nhật thông tin cho record đã tồn tại
		if (StringUtils.isNotBlank(curBacklogKeyVal) && curBacklog != null) {
			final var wrRemoveEles = fillDataForRow(sheet, curIdx, curBacklog, allWrDatas, indexNo);

			allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule
		}

		// Điền thông tin cho record được thêm mới
		var stepCnt = 0;
		for (final BacklogDetail backlogDetail : newBacklogs) {

			final var curRowIdx = startRowShift - parentRowCnt + stepCnt;

			final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas, indexNo);

			allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

			stepCnt++;
		}

		// Merge lại cell và điền dữ liệu
		if (!newBacklogs.isEmpty()) {
			final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
			final var ankenNoCell = row.getCell(COLUMN_ANKEN_INDEX);
			formulaEvaluator.evaluate(ankenNoCell);
			final var fRow = curIdx;
			final var lRow = fRow + numberOfRowsToShift;
			// merge cell
			// Column No
			var newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_A_INDEX, COLUMN_A_INDEX);
			sheet.addMergedRegion(newMergedRegion);

			// Column "グループ Group"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_B_INDEX, COLUMN_B_INDEX);
			sheet.addMergedRegion(newMergedRegion);

			// Column "画面ID Screen ID"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_ANKEN_INDEX, COLUMN_ANKEN_INDEX);
			sheet.addMergedRegion(newMergedRegion);

			// Column "画面名 Screen Name"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_SCREEN_INDEX, COLUMN_SCREEN_INDEX);
			sheet.addMergedRegion(newMergedRegion);

			// Column "ステータス Status"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_STATUS_INDEX, COLUMN_STATUS_INDEX);
			sheet.addMergedRegion(newMergedRegion);
		}
		return numberOfRowsToShift + curRowCnt - 1;
	}

	private int fillWhenExistsMutiRecord(final Sheet sheet, final Row row, final int curIdx,
			final AtomicInteger indexNo, final CellRangeAddress mergeCellRange, final List<BacklogDetail> backlogs,
			final List<PjjyujiDetail> allWrDatas) {

		// Lấy phạm vi của MergeCell
		final var firstRowIdx = mergeCellRange.getFirstRow();
		final var lastRowIdx = mergeCellRange.getLastRow();

		var from = firstRowIdx; // Starting number
		final var to = lastRowIdx; // Ending number
		final List<String> backLogKeyExists = new ArrayList<>();
		while (from <= to) {
			final var curBacklogKey = getBacklogKeyVal(sheet.getRow(from));
			backLogKeyExists.add(curBacklogKey);
			from++;
		}
		final var newBacklogs = backlogs.stream()
				.filter(x -> backLogKeyExists.stream().allMatch(k -> !StringUtils.equals(k, x.getKey()))).toList();

		final var curBacklogs = backlogs.stream()
				.filter(x -> backLogKeyExists.stream().anyMatch(k -> StringUtils.equals(k, x.getKey()))).toList();

		final var curRowCnt = backLogKeyExists.size();
		// Dịch chuyển các dòng
		final var startRowShift = curIdx + curRowCnt;
		final var numberOfRowsToShift = newBacklogs.size();
		if (numberOfRowsToShift > 0) {
			shiftRow(sheet, startRowShift, numberOfRowsToShift);

			// Tạo dòng mới sau khi dịch chuyển
			for (var j = startRowShift; j < startRowShift + numberOfRowsToShift; j++) {
				final var newRow = sheet.createRow(j);
				cloneRowFormat(row, newRow);
			}
		}

		// Cập nhật thông tin cho record đã tồn tại
		var stepCnt = 0;
		for (final BacklogDetail backlogDetail : curBacklogs) {

			final var curRowIdx = curIdx + stepCnt;

			final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas, indexNo);

			allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

			stepCnt++;
		}
		stepCnt = 0;
		// Điền thông tin cho record được thêm mới
		for (final BacklogDetail backlogDetail : newBacklogs) {

			final var curRowIdx = startRowShift + stepCnt;

			final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas, indexNo);

			allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

			stepCnt++;
		}
		if (!newBacklogs.isEmpty()) {
			final var fRow = curIdx;
			final var lRow = fRow + curRowCnt + numberOfRowsToShift - 1;

			// unmerge cell
			for (var k = sheet.getNumMergedRegions() - 1; k >= 0; k--) {
				final var mergedRegion = sheet.getMergedRegion(k);
				if (compareCellRangeAddresses(mergedRegion, mergeCellRange)) {
					sheet.removeMergedRegion(k);
				}
			}

			// merge cell
			// Column No
			var newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_A_INDEX, COLUMN_A_INDEX);
			sheet.addMergedRegion(newMergedRegion);
			// Column "グループ Group"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_B_INDEX, COLUMN_B_INDEX);
			sheet.addMergedRegion(newMergedRegion);

			// Column "画面ID Screen ID"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_ANKEN_INDEX, COLUMN_ANKEN_INDEX);
			sheet.addMergedRegion(newMergedRegion);

			// Column "画面名 Screen Name"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_SCREEN_INDEX, COLUMN_SCREEN_INDEX);
			sheet.addMergedRegion(newMergedRegion);
			// Column "ステータス Status"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_STATUS_INDEX, COLUMN_STATUS_INDEX);
			sheet.addMergedRegion(newMergedRegion);
		}
		return numberOfRowsToShift + curRowCnt - 1;
	}

	private int fillWhenNewRecord(final Sheet sheet, final Row row, final int curIdx, final AtomicInteger indexNo,
			final String curBacklogParentKey, final List<BacklogDetail> backlogs,
			final List<PjjyujiDetail> allWrDatas) {

		// Lấy ra ticket no
		final var curAnkenNo = backlogs.stream().findFirst().map(BacklogDetail::getAnkenNo).orElse("");
		final Integer totalRow = backlogs.size();
//			row.getCell(columnAIndex).setCellValue(curIdx + 1); // fill number no

		// Dịch chuyển các dòng
		final var startRowShift = curIdx + 1;
		final int numberOfRowsToShift = totalRow;
		if (numberOfRowsToShift > 0) {
			shiftRow(sheet, startRowShift, numberOfRowsToShift);
			// Tạo dòng mới sau khi dịch chuyển
			for (var j = startRowShift; j < startRowShift + numberOfRowsToShift; j++) {
				final var newRow = sheet.createRow(j);
				cloneRowFormat(row, newRow);
			}
		}

		// Điền dữ liệu cho các dòng thêm mới
		var stepCnt = 0;
		for (final BacklogDetail backlogDetail : backlogs) {

			final var curRowIdx = curIdx + stepCnt;

			final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas, indexNo);

			allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

			stepCnt++;
		}

		if (totalRow > 1) {
			final var fRow = curIdx;
			final var lRow = fRow + numberOfRowsToShift - 1;

			// merge cell
			// Column No
			var newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_A_INDEX, COLUMN_A_INDEX);
			sheet.addMergedRegion(newMergedRegion);
			// Column "グループ Group"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_B_INDEX, COLUMN_B_INDEX);
			sheet.addMergedRegion(newMergedRegion);
			setValForMergeCell(sheet, newMergedRegion, COLUMN_B_INDEX, curBacklogParentKey);

			// Column "画面ID Screen ID"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_ANKEN_INDEX, COLUMN_ANKEN_INDEX);
			sheet.addMergedRegion(newMergedRegion);
			setValForMergeCell(sheet, newMergedRegion, COLUMN_ANKEN_INDEX, curAnkenNo);

			// Column "画面名 Screen Name"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_SCREEN_INDEX, COLUMN_SCREEN_INDEX);
			sheet.addMergedRegion(newMergedRegion);
			// Column "ステータス Status"
			newMergedRegion = new CellRangeAddress(fRow, lRow, COLUMN_STATUS_INDEX, COLUMN_STATUS_INDEX);
			sheet.addMergedRegion(newMergedRegion);
		}
		return numberOfRowsToShift - 1;
	}

	private int fillOnlyParentRecord(final Sheet sheet, final Row row, final int curIdx, final AtomicInteger indexNo,
			final Optional<BacklogDetail> parentBacklog, final List<PjjyujiDetail> wrTargets) {
		// Lấy ra ticket no
		parentBacklog.ifPresent(b -> {
			final Integer totalRow = 1;
			// Dịch chuyển các dòng
			final var startRowShift = curIdx + 1;
			final int numberOfRowsToShift = totalRow;
			if (numberOfRowsToShift > 0) {
				shiftRow(sheet, startRowShift, numberOfRowsToShift);
				// Tạo dòng mới sau khi dịch chuyển
				for (var j = startRowShift; j < startRowShift + numberOfRowsToShift; j++) {
					final var newRow = sheet.createRow(j);
					cloneRowFormat(row, newRow);
				}
			}
			fillDataForRow(sheet, curIdx, b, wrTargets, indexNo);
		});
		return 0;
	}

	private void fillDataForSheet(final Workbook workbook, final Sheet sheet, final List<BacklogDetail> allBacklogs,
			final List<PjjyujiDetail> allWrDatas) {
		if (CollectionUtils.isEmpty(allBacklogs)) {
			return;
		}
		final var groupedBacklogs = groupByParentKey(allBacklogs);
		final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		formulaEvaluator.evaluateAll();
		final var dataFormatter = new DataFormatter();

		// Điền thông tin chung
		fillCommonInfo(sheet, groupedBacklogs, allWrDatas);

		final var curRowIdxAtomic = new AtomicInteger(0);
		final var indexNo = new AtomicInteger(1);
		while (MapUtils.isNotEmpty(groupedBacklogs)) {
			final var curIdx = curRowIdxAtomic.getAndIncrement();
			var row = sheet.getRow(curIdx);
			// skip xử lý khi đang đọc các dòng header
			if (curIdx <= TARGET_DATE_ROW_IDX) {
				continue;
			}
			if (row == null) {
				row = sheet.createRow(curIdx);
				cloneRowFormat(sheet.getRow(curIdx - 1), row);
			}
			log.debug("Sheet {} at row {}; remaining group {}", sheet.getSheetName(), curIdx, groupedBacklogs.size());

			final var groupCell = row.getCell(COLUMN_B_INDEX);
			var curBacklogParentKey = dataFormatter.formatCellValue(groupCell, formulaEvaluator);
			final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(groupCell);
			final var isSingleRecord = mergeCellRange == null;
			if (!isSingleRecord) {
				curBacklogParentKey = StringUtils.trim(ScheduleHelper.readContentCell(sheet, groupCell));
			}
			final var isExists = StringUtils.isNotBlank(curBacklogParentKey)
					&& groupedBacklogs.containsKey(curBacklogParentKey);

			if (!isSingleRecord && !isExists) {
				// Skip record dang ton tai nhung khong co thong tin update
				// Lấy phạm vi của MergeCell
				final var firstRowIdx = mergeCellRange.getFirstRow();
				final var lastRowIdx = mergeCellRange.getLastRow();
				final var increCnt = lastRowIdx - firstRowIdx;
				curRowIdxAtomic.addAndGet(increCnt);
				indexNo.incrementAndGet();
				continue;
			}
			// T/h tồn tại thực hiện cập nhật thông tin, thêm dòng mới, merge cell lại
			if (isExists) {
				var increCnt = 0;
				final var backlogs = getTargetBacklogs(groupedBacklogs, curBacklogParentKey);
				if (isSingleRecord) {
					increCnt = fillWhenExistsSingleRecord(sheet, row, curIdx, indexNo, backlogs, allWrDatas);
				} else {
					increCnt = fillWhenExistsMutiRecord(sheet, row, curIdx, indexNo, mergeCellRange, backlogs,
							allWrDatas);
				}
				// tăng index xử lý sau khi xử lý thêm dòng
				curRowIdxAtomic.addAndGet(increCnt);
			} else {
				// T/h new mới row
				final var firstEntryOptional = groupedBacklogs.entrySet().stream().findFirst();
				if (!firstEntryOptional.isPresent()) {
					break; // exit when no record exists
				}
				final var firstEntry = firstEntryOptional.get();
				curBacklogParentKey = firstEntry.getKey(); // Lấy ra parent key
				final var curBacklogs = firstEntry.getValue();
				final var parentBacklog = curBacklogs.stream().filter(isBacklogParent).findFirst();
				final var backlogs = curBacklogs.stream().filter(isBacklogDetail).toList();

				var increCnt = 0;
				if (backlogs.isEmpty() && parentBacklog.isPresent()) {
					increCnt = fillOnlyParentRecord(sheet, row, curIdx, indexNo, parentBacklog, allWrDatas);
				} else {
					increCnt = fillWhenNewRecord(sheet, row, curIdx, indexNo, curBacklogParentKey, backlogs,
							allWrDatas);
				}
				curRowIdxAtomic.addAndGet(increCnt);
			}
			// remove sau khi lay ra thong tin xu ly
			groupedBacklogs.remove(curBacklogParentKey);
			indexNo.incrementAndGet();
		}

		fillTotalPic(sheet, allWrDatas);

		evaluate(workbook, sheet);

	}

	private void fillTotalPic(final Sheet sheet, final List<PjjyujiDetail> allWrDatas) {
		for (final Row row : sheet) {
			if (isTotalRow(row)) {
				final var pics = allWrDatas.stream().map(PjjyujiDetail::getMailId).distinct().sorted()
						.collect(Collectors.toList());
				final var aiIdx = new AtomicInteger(row.getRowNum());
				while (CollectionUtils.isNotEmpty(pics)) {
					final var targetCell = sheet.getRow(aiIdx.incrementAndGet())
							.getCell(CellReference.convertColStringToIndex(COLUMN_TOTAL_CHARACTER));
					final var val = ScheduleHelper.getCellValueAsString(targetCell);
					if (StringUtils.isNotBlank(val)) {
						if (!pics.contains(val)) {
							continue;
						}
						pics.removeIf(x -> StringUtils.equals(x, val));
					} else {
						targetCell.setCellValue(pics.get(0));
						pics.remove(0);
					}
				}
				break;
			}
		}
	}

	private void evaluate(final Workbook workbook, final Sheet sheet) {
		// Cập nhật lại công thức
		updatedTotalActualHoursFormula(sheet);

		// Chạy lại toàn bộ công thức
		evaluateAllFormula(workbook);
	}

	private String extracProcessOfWrCd(final String input) {
		final var pattern = "(\\d+):";

		final var regex = Pattern.compile(pattern);
		final var matcher = regex.matcher(input);

		if (matcher.find()) {
			return matcher.group(1);
		}
		return WorkingPhase.ID0.getCode();
	}

	private final Predicate<BacklogDetail> isBacklogBug = backlogDetail -> {
		final var processOfWrCd = extracProcessOfWrCd(backlogDetail.getProcessOfWr());
		final var wrBugCd = WorkingPhase.ID43.getCode();
		return StringUtils.equals(wrBugCd, processOfWrCd);
	};

	private final Predicate<BacklogDetail> isBacklogSpec = backlogDetail -> {
		final var processOfWrCd = extracProcessOfWrCd(backlogDetail.getProcessOfWr());
		final var wrBugCd = WorkingPhase.ID45.getCode();
		return StringUtils.equals(wrBugCd, processOfWrCd);
	};

	private final Predicate<BacklogDetail> isBacklogPg = backlogDetail -> isBacklogBug.negate()
			.and(isBacklogSpec.negate()).test(backlogDetail);

	private void fillBacklogDetailInfo(final Workbook workbook, final List<BacklogDetail> bds,
			final List<PjjyujiDetail> pds) {

		final var backlogBug = bds.stream() //
				.filter(isBacklogBug) //
				.toList();
		final var backlogSpec = bds.stream() //
				.filter(isBacklogSpec) //
				.toList();
		final var backlogPg = bds.stream() //
				.filter(isBacklogPg) //
				.toList();
		log.debug("fillBacklogInfo.PG->CNT: {}", backlogPg.size());
		log.debug("fillBacklogInfo.SPEC->CNT: {}", backlogSpec.size());
		log.debug("fillBacklogInfo.BUG->CNT: {}", backlogBug.size());
		final var sheetIterator = workbook.sheetIterator();
		while (sheetIterator.hasNext()) {
			final var sheet = sheetIterator.next();
			// Kiểm tra là sheet điền schedule
			if (!ScheduleHelper.isScheduleSheet(sheet)) {
				continue;
			}
			genScheduleInfoForSheet(workbook, sheet, backlogPg, backlogSpec, backlogBug, pds);
		}
	}

	private void genScheduleInfoForSheet(final Workbook workbook, final Sheet sheet, final List<BacklogDetail> pgs,
			final List<BacklogDetail> specs, final List<BacklogDetail> bugs, final List<PjjyujiDetail> pds) {

		final var sheetName = StringUtils.lowerCase(sheet.getSheetName());
		log.debug("fillBacklogInfo.sheetName: {}", sheetName);

		var datas = pgs;
		if (StringUtils.equals(sheetName, "pg_spec")) {
			datas = specs;
		} else if (StringUtils.equals(sheetName, "pg_bug")) {
			datas = bugs;
		}

		fillDataForSheet(workbook, sheet, datas, pds);

		// Cập nhật lại công thức
		updatedTotalActualHoursFormula(sheet);

		// Chạy lại toàn bộ công thức
		evaluateAllFormula(workbook);
	}

	private Pair<YearMonth, YearMonth> getRangeTarget(final List<PjjyujiDetail> pds) {
		final var yearMonths = pds.stream().map(PjjyujiDetail::getTargetYmd).map(YearMonth::from).distinct().toList();
		final var now = YearMonth.now();
		final var ymS = yearMonths.stream().min(YearMonth::compareTo).orElse(now);
		final var ymE = yearMonths.stream().max(YearMonth::compareTo).orElse(now);
		return Pair.of(ymS, ymE);
	}

	private void createRangeWorkingReportDetail(final List<PjjyujiDetail> pds, final Workbook workbook) {

		final var yearMonths = getRangeTarget(pds);
		final var ymS = yearMonths.getLeft();
		final var ymE = yearMonths.getRight();

		final var sheetIterator = workbook.sheetIterator();
		while (sheetIterator.hasNext()) {
			final var sheet = sheetIterator.next();
			if (!ScheduleHelper.isScheduleSheet(sheet)) {
				continue;
			}
			standardizedRangeInput(sheet, ymS, ymE);
		}
		evaluateAllFormula(workbook);
	}

	/**
	 *
	 * @param pds
	 * @param bds
	 * @param workbook
	 */
	private void fillScheduleInfo(final List<PjjyujiDetail> pds, final List<BacklogDetail> bds,
			final Workbook workbook) {

		createRangeWorkingReportDetail(pds, workbook);

		fillBacklogDetailInfo(workbook, bds, pds);

		// Chạy lại toàn bộ công thức
		evaluateAllFormula(workbook);
	}

	private Path getLastSchedule(final Path projectSchPath) throws IOException {
		final var filePattern = "QDA-0222a_プロジェクト管理表_(\\d{8})_(\\d{6})(?:_(\\d{6}))?\\.xlsm";

		try (var directoryStream = Files.newDirectoryStream(projectSchPath)) {
			Path lastFile = null;
			var lastEndDate = YearMonth.from(LocalDate.MIN); // Initialize to a very early date
			final var pattern = Pattern.compile(filePattern, Pattern.CANON_EQ);
			final var dateFormatter = DateTimeFormatter.ofPattern("yyyyMM");

			for (final Path file : directoryStream) {
				if (Files.isRegularFile(file)) {
					final var fileName = file.getFileName().toString();
					final var matcher = pattern.matcher(fileName);

					if (matcher.find()) {
//						final var group1 = matcher.group(1); // projectCd
						final var group2 = matcher.group(2); // start Year month
						final var group3 = matcher.group(3); // end Year month
						final var endYearMonthString = StringUtils.isNoneBlank(group3) ? group3 : group2;
						final var endDate = YearMonth.parse(endYearMonthString, dateFormatter);

						if (endDate.isAfter(lastEndDate)) {
							lastFile = file;
							lastEndDate = endDate;
						}
					}
				}
			}
			return lastFile;
		}
	}

	private Path createFolderStoreSchedule(final CustomerTarget projecType, final String projectCd) throws IOException {

		final var projectScheduleTemplate = switch (projecType) {
		case IFRONT -> PATH_IFRONT_TEMPLATE;
		case SYMPHONIZER -> PATH_SYM_TEMPLATE;
		case DMP -> PATH_DMP_TEMPLATE;
		default -> PATH_DEFAULT_TEMPLATE;
		};
		final var projectSchPath = Paths
				.get(String.format(projectScheduleTemplate, PATH_ROOT_FOLDER, executeTime, projectCd));
		if (!Files.exists(projectSchPath)) {
			Files.createDirectories(projectSchPath);
		}
		return projectSchPath;
	}

	public void createSchedule(final CustomerTarget projecType, final String projectCd, final List<PjjyujiDetail> pds,
			final List<BacklogDetail> bds) throws IOException {
		log.debug("Bat dau tao schedule: {} - {}", projecType, projectCd);

		final var projectSchPath = createFolderStoreSchedule(projecType, projectCd);

		var isUpdateSchedule = false;
		Path lastSchePath = null;
		if (isUpdateOldSchedule()) {
			lastSchePath = getLastSchedule(projectSchPath);
			isUpdateSchedule = isUpdateOldSchedule() && lastSchePath != null;
			if (isUpdateSchedule) {
				ZipSecureFile.setMinInflateRatio(0);
			}
		}

		try (var fis = isUpdateSchedule ? new FileInputStream(lastSchePath.toFile())
				: BacklogExcel.class.getClassLoader().getResourceAsStream(SCHEDULE_TEMPLATE_PATH);
				Workbook workbook = new XSSFWorkbook(fis)) {

			fillScheduleInfo(pds, bds, workbook);

			// new file schedule
			final var targetFile = createNewFileSchedule(projectCd, projectSchPath, pds);

			// ghi vào file schedule mới
			final var schFilePath = saveToNewFileSchedule(workbook, targetFile);

			log.debug("Ket thuc tao schedule: {} - {}", projectCd, schFilePath);
		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

	private File createNewFileSchedule(final String projecCd, final Path projectSchPath,
			final List<PjjyujiDetail> pds) {

		final var yearMonths = getRangeTarget(pds);
		final var ymS = yearMonths.getLeft();
		final var ymE = yearMonths.getRight();
		var sufFileName = "";

		// Define the desired format
		final var formatter = DateTimeFormatter.ofPattern("yyyyMM");

		if (ymS.compareTo(ymE) == 0) {
			sufFileName = ymS.format(formatter);
		} else {
			sufFileName = ymS.format(formatter) + "_" + ymE.format(formatter);
		}
		// Ghi dữ liệu vào tệp tin
		final var fileName = StringUtils.replaceEach(TEMPLATE_FILE, new String[] { "{projectCd}", "{range}" },
				new String[] { projecCd, sufFileName });
		// new file schedule
		File targetFile = null;
		if (projectSchPath != null) {
			final var filePath = projectSchPath.resolve(fileName);
			// Convert the Path object to a File object
			targetFile = filePath.toFile();
		} else {
			targetFile = new File(fileName);
		}
		return targetFile;
	}

	/**
	 * Sau khi xử lý xong schedule thì thực hiện ghi vào file mới
	 *
	 * @param workbook
	 * @param pjCd
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private String saveToNewFileSchedule(final Workbook workbook, final File targetFile) throws IOException {

		try (var fileOut = new FileOutputStream(targetFile, false);) {

			workbook.write(fileOut);
		}
		return targetFile.getAbsolutePath();
	}

	private void evaluateAllFormula(final Workbook workbook) {
		// Create a formula evaluator
		final var evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		// Update all formulas in the sheet
		evaluator.evaluateAll();

	}

	@Override
	public boolean isUpdateOldSchedule() {
		return false;
	}

	private Map<String, List<BacklogDetail>> groupByParentKey(final List<BacklogDetail> details) {
		final Map<String, List<BacklogDetail>> groupedBacklogs = new HashMap<>();

		for (final BacklogDetail obj : details) {
			if (StringUtils.isNotBlank(obj.getParentKey())) {
				groupedBacklogs.computeIfAbsent(obj.getParentKey(), key -> new ArrayList<>()).add(obj);
			} else {
				groupedBacklogs.computeIfAbsent(obj.getKey(), key -> new ArrayList<>()).add(obj);
			}
		}
		return groupedBacklogs;
	}

	@Override
	public boolean isExportZip() {
		return true;
	}

}

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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.NumberToTextConverter;
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

	private static final String columnACharacter = "A";
	private static final String columnBCharacter = "B";
	private static final String columnAnkenCharacter = "C";
	private static final String columnScreenCharacter = "D";
	private static final String colOperationChar = "E";
	private static final String columnPicCharacter = "F";
	private static final String columnTotalCharacter = "F";
	private static final String columnStatusCharacter = "G";
	private static final String colExpectHousrChar = "H";
	private static final String colExpectStartYmdChar = "I";
	private static final String colExpectEndYmdChar = "J";
	private static final String colExpectDeliveryYmdChar = "K";
	private static final String colBacklogIdChar = "L";
	private static final String colActualTotalHoursBacklogChar = "M";
	private static final String colActHousrChar = "N";
	private static final String colActStartYmdChar = "O";
	private static final String colActEndYmdChar = "P";
	private static final String colActProgressChar = "Q";
	private static final String colActDeliveryYmdChar = "R";
	private static final String colTemplateStartDate = "S";

	private static final String totalCharacter = "Total";
	private static final int columnAIndex = CellReference.convertColStringToIndex(columnACharacter);
	private static final int columnBIndex = CellReference.convertColStringToIndex(columnBCharacter);
	private static final int columnAnkenIndex = CellReference.convertColStringToIndex(columnAnkenCharacter);
	private static final int columnScreenIndex = CellReference.convertColStringToIndex(columnScreenCharacter);
	private static final int columnStatusIndex = CellReference.convertColStringToIndex(columnStatusCharacter);
	private static final int targetMonthRowIdx = 7;
	private static final int targetDateRowIdx = 8;
	private static final int columnStartDateInputIdx = CellReference.convertColStringToIndex(colTemplateStartDate);
	private static final DateTimeFormatter FORMATTER_YYYYMMDD = DateTimeFormatter.ofPattern("yyyy/MM/dd");

	private static final String templateFile = "QDA-0222a_プロジェクト管理表_{projectCd}_{range}.xlsm";
	private static final String templateTotalActHours = "SUM({cNameS}{rIdx}:{cNameE}{rIdx})";
	private static final String templateNextDateFormula = "{preCol}+1";

	private static final String scheduleTemplatePath = "templates/QDA-0222a_プロジェクト管理表.xlsm";

	public static final String pathRootFolder = "D:\\Doc\\Backlog";
	private static final String pathSymTemplate = "%s\\sym\\%s";
	private static final String pathIfrontTemplate = "%s\\ifront\\%s";
	private static final String pathDefaultTemplate = "%s\\default\\%s";

	private boolean compareCellRangeAddresses(final CellRangeAddress range1, final CellRangeAddress range2) {
		// Compare the first row, last row, first column, and last column
		return range1.getFirstRow() == range2.getFirstRow() && range1.getLastRow() == range2.getLastRow();
	}

	private boolean isTotalRow(final Row row) {
		final var columnTotalIndex = CellReference.convertColStringToIndex(columnTotalCharacter);
		return ScheduleHelper.isTotalRow(row, totalCharacter, columnTotalIndex);
	}

	/*
	 * Cập nhật công thức tính tổng dựa trên việc thêm cột mới.
	 */
	private void updatedTotalActualHoursFormula(final Sheet sheet) {
		final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var columnIndex = CellReference.convertColStringToIndex(colActHousrChar);
		for (final Row row : sheet) {
			if (isTotalRow(row)) {
				break;
			}
			final var chr = ScheduleHelper.convertColumnIndexToName(row.getLastCellNum());
			final var rNum = row.getRowNum();
			if (rNum >= 9) {
				final var cell = row.getCell(columnIndex);
				if (cell != null) {
					final var adjustedFormula = StringUtils.replaceEach(templateTotalActHours,
							new String[] { "{rIdx}", "{cNameS}", "{cNameE}" },
							new String[] { String.valueOf(rNum + 1), colTemplateStartDate, chr });
					cell.setCellFormula(adjustedFormula);
					formulaEvaluator.evaluate(cell);
				}
			}
		}
	}

	/**
	 * Cập nhật công thức cho vùng footer total
	 *
	 * @param sheet
	 * @param formulaEvaluator
	 */
	private void updatedTotalFooterFormula(final Sheet sheet) {
		// TODO
	}

	public String getCellValue(final Cell cell) {
		if (cell == null) {
			return "";
		}

		String cellValue;
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				cellValue = cell.getDateCellValue().toString();
			} else {
				cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
			}
			break;
		case BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case FORMULA:
			cellValue = evaluateFormulaCell(cell);
			break;
		case BLANK:
			cellValue = "";
			break;
		default:
			cellValue = "";
		}

		return cellValue;
	}

	private String evaluateFormulaCell(final Cell cell) {
		final var formulaEvaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var cellValue = formulaEvaluator.evaluate(cell);
		return getCellValueFromFormulaResult(cellValue);
	}

	private String getCellValueFromFormulaResult(final CellValue cellValue) {
		return switch (cellValue.getCellType()) {
		case STRING -> cellValue.getStringValue();
		case NUMERIC -> NumberToTextConverter.toText(cellValue.getNumberValue());
		case BOOLEAN -> String.valueOf(cellValue.getBooleanValue());
		default -> "";
		};
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
			}
		}
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
		return row.getRowNum() <= targetDateRowIdx;
	}

	private Date toDate(final LocalDate localDate) {
		if (localDate == null) {
			return null;
		}
		// Convert LocalDate to Date
		final var localDateTime = localDate.atStartOfDay();
		final var zonedDateTime = localDateTime.atZone(ZoneId.systemDefault());
		final var date = Date.from(zonedDateTime.toInstant());
		return date;
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
		var curCel = getCell(curRow, columnACharacter);
		curCel.setCellValue(indexNo.get());
		// "グループ Group"
		final var parentKey = Optional.ofNullable(backlogDetail).map(BacklogDetail::getParentKey)
				.orElse(StringUtils.EMPTY);
		curCel = getCell(curRow, columnBCharacter);
		curCel.setCellValue(parentKey);
		// "画面ID Screen ID"
		final var ankenNo = Optional.ofNullable(backlogDetail).map(BacklogDetail::getAnkenNo).orElse(StringUtils.EMPTY);
		curCel = getCell(curRow, columnAnkenCharacter);
		curCel.setCellValue(ankenNo);
		// 工程 Operation
		final var operation = getOperation(backlogDetail).map(WorkingPhase::getName).orElse(StringUtils.EMPTY);
		curCel = getCell(curRow, colOperationChar);
		curCel.setCellValue(operation);
		// 担当 PIC
		final var pic = backlogDetail.getMailId();
		curCel = getCell(curRow, columnPicCharacter);
		curCel.setCellValue(pic);
		// ステータス Status
		// curCel =
		// curRow.getCell(CellReference.convertColStringToIndex(columnStatusCharacter));
		// curCel.setCellValue(backlogDetail.getStatus());

		// "予定 Schedule"
		// 工数 Hours
		curCel = getCell(curRow, colExpectHousrChar);
		curCel.setCellValue(
				Optional.ofNullable(backlogDetail.getEstimatedHours()).orElse(BigDecimal.ZERO).doubleValue());
		// 開始日 Begin
		curCel = getCell(curRow, colExpectStartYmdChar);
		curCel.setCellValue(toDate(backlogDetail.getExpectedStartDate()));
		// 完了日 End
		curCel = getCell(curRow, colExpectEndYmdChar);
		curCel.setCellValue(toDate(backlogDetail.getExpectedDueDate()));
		// 納品日 Delivery
		curCel = getCell(curRow, colExpectDeliveryYmdChar);
		curCel.setCellValue(toDate(backlogDetail.getExpectedDeliveryDate()));

		// Backlog Information
		// Key
		curCel = getCell(curRow, colBacklogIdChar);
		curCel.setCellValue(backlogDetail.getKey());
		// Hours
		curCel = getCell(curRow, colActualTotalHoursBacklogChar);
		curCel.setCellValue(Optional.ofNullable(backlogDetail.getActualHours()).orElse(BigDecimal.ZERO).doubleValue());

		// "実績 Actual"
		// 開始日 Begin
		curCel = getCell(curRow, colActStartYmdChar);
		curCel.setCellValue(toDate(backlogDetail.getActualStartDate()));
		// 完了日 End
		curCel = getCell(curRow, colActEndYmdChar);
		curCel.setCellValue(toDate(backlogDetail.getActualDueDate()));
		// 進捗 Progress
		curCel = getCell(curRow, colActProgressChar);
		curCel.setCellValue(getProgress(backlogDetail.getProgress()));
		// 納品日 Delivery
		curCel = getCell(curRow, colActDeliveryYmdChar);
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
			final var targetDateRow = sheet.getRow(targetDateRowIdx);
			final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
			for (final Cell c : curRow) {
				final var curColIdx = c.getColumnIndex();
				if (curColIdx < columnStartDateInputIdx) {
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
		final var row = sheet.getRow(targetMonthRowIdx);
		String lastTarget = null;
		for (final Cell cell : row) {
			if (cell == null || cell.getColumnIndex() < columnStartDateInputIdx) {
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

	private void addColInput(final Sheet sheet, final YearMonth targetYm, final boolean isFirstTargetMonth) {
		final var row = sheet.getRow(targetMonthRowIdx);
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
		for (var i = targetMonthRowIdx; i <= sheet.getLastRowNum(); i++) {
			final var sourceRow = sheet.getRow(i);
			var destinationRow = sheet.getRow(i);
			if (sourceRow != null) {
				final var sourceCell = sourceRow.getCell(lastColumnIndex);
				if (sourceCell != null && i <= targetDateRowIdx && isFirstTargetMonth) { // first target month
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
					if (!isFirstTargetMonth && j == 1 && i <= targetDateRowIdx) {
						destinationCell.setCellValue(localDate);
					}
					// fill formula plus date
					if (i == targetDateRowIdx) {
						final var adjustedFormula = StringUtils.replaceEach(templateNextDateFormula,
								new String[] { "{preCol}", }, new String[] { preColStr + (targetDateRowIdx + 1) });
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
				if (i == targetMonthRowIdx) {
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
		final var obj = new BacklogExcel();
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

		final var backlogKeyVal = StringUtils.trim(StringUtils.defaultString(dataFormatter.formatCellValue(
				row.getCell(CellReference.convertColStringToIndex(colBacklogIdChar)), formulaEvaluator)));
		return backlogKeyVal;
	}

	private void shiftRow(final Sheet sheet, final int startRowShift, final int numberOfRowsToShift) {
		if (startRowShift <= sheet.getLastRowNum()) {
			sheet.shiftRows(startRowShift, sheet.getLastRowNum(), numberOfRowsToShift);
		} else {
			log.debug("Sheet {} curLastRowNum {} startRowShift {} numberOfRowsToShift {}", sheet.getSheetName(),
					startRowShift, numberOfRowsToShift);
//			sheet.shiftRows(startRowShift, startRowShift, numberOfRowsToShift);
//			curLastRowNum.getAndAdd(numberOfRowsToShift);
		}
	}

	private void fillDataForSheet(final Workbook workbook, final Sheet sheet, final List<BacklogDetail> allBacklogs,
			final List<PjjyujiDetail> allWrDatas) {
		if (CollectionUtils.isEmpty(allBacklogs)) {
			return;
		}
		final var groupedBacklogs = groupByParentKey(allBacklogs);
		final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		final var dataFormatter = new DataFormatter();

		final var i = new AtomicInteger(0);
		final var indexNo = new AtomicInteger(1);
		while (MapUtils.isNotEmpty(groupedBacklogs)) {
			sheet.getSheetName();
			final var curIdx = i.getAndIncrement();
			var row = sheet.getRow(curIdx);
			// skip xử lý khi đang đọc các dòng header
			if (curIdx <= targetDateRowIdx) {
				continue;
			}
			if (row == null) {
				row = sheet.createRow(curIdx);
				cloneRowFormat(sheet.getRow(curIdx - 1), row);
			}

//			log.debug("Sheet {} curLastRowNum {}", sheet.getSheetName(), sheet.getLastRowNum());
			log.debug("Sheet {} at row {} - {}", sheet.getSheetName(), curIdx, groupedBacklogs.size());

			final var groupCell = row.getCell(columnBIndex);
			formulaEvaluator.evaluate(groupCell);
			var curBacklogParentKey = dataFormatter.formatCellValue(groupCell, formulaEvaluator);
			final var isExists = StringUtils.isNotBlank(curBacklogParentKey)
					&& groupedBacklogs.containsKey(curBacklogParentKey);

			final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(groupCell);
			final var isSingleRecord = mergeCellRange == null;
			var numberOfRowsToShift = 0;
			// T/h tồn tại thực hiện cập nhật thông tin, thêm dòng mới, merge cell lại
			if (isExists) {
				var curRowCnt = 0;
				final var backlogs = CollectionUtils.emptyIfNull(groupedBacklogs.get(curBacklogParentKey));
				// update atomic value after fill
				if (isSingleRecord) {
					curRowCnt = 1;
					final var curBacklogKeyVal = getBacklogKeyVal(row);
					final var newBacklogs = backlogs.stream()
							.filter(x -> !StringUtils.equals(x.getKey(), curBacklogKeyVal))
							.collect(Collectors.toList());
					final var curBacklog = backlogs.stream()
							.filter(x -> StringUtils.equals(x.getKey(), curBacklogKeyVal)).findFirst().orElse(null);

					// Dịch chuyển các dòng
					final var startRowShift = curIdx + curRowCnt;
					numberOfRowsToShift = newBacklogs.size();
					if (numberOfRowsToShift > 0) {
						shiftRow(sheet, startRowShift, numberOfRowsToShift);

						// Tạo dòng mới sau khi dịch chuyển
						for (var j = startRowShift; j < startRowShift + numberOfRowsToShift; j++) {
							final var newRow = sheet.createRow(j);
							cloneRowFormat(row, newRow);
						}
					}

					// Cập nhật thông tin cho record đã tồn tại
					if (curBacklog != null) {
						final var wrRemoveEles = fillDataForRow(sheet, curIdx, curBacklog, allWrDatas, indexNo);

						allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule
					}

					// Điền thông tin cho record được thêm mới
					var stepCnt = 0;
					for (final BacklogDetail backlogDetail : newBacklogs) {

						final var curRowIdx = startRowShift + stepCnt;

						final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas, indexNo);

						allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

						stepCnt++;
					}

					// Merge lại cell và điền dữ liệu
					if (!newBacklogs.isEmpty()) {
						final var ankenNoCell = row.getCell(columnAnkenIndex);
						formulaEvaluator.evaluate(ankenNoCell);
//							final var curAnkenNo = dataFormatter.formatCellValue(ankenNoCell, formulaEvaluator);
						final var fRow = curIdx;
						final var lRow = fRow + numberOfRowsToShift;
						// merge cell
						// Column No
						var newMergedRegion = new CellRangeAddress(fRow, lRow, columnAIndex, columnAIndex);
						sheet.addMergedRegion(newMergedRegion);
						// Column "グループ Group"
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnBIndex, columnBIndex);
						sheet.addMergedRegion(newMergedRegion);
//							setValForMergeCell(sheet, newMergedRegion, columnBIndex, curBacklogParentKey);

						// Column "画面ID Screen ID"
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnAnkenIndex, columnAnkenIndex);
						sheet.addMergedRegion(newMergedRegion);
//							setValForMergeCell(sheet, newMergedRegion, columnAnkenIndex, curAnkenNo);

						// Column "画面名 Screen Name"
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex, columnScreenIndex);
						sheet.addMergedRegion(newMergedRegion);
						// Column "ステータス Status"
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex, columnStatusIndex);
						sheet.addMergedRegion(newMergedRegion);
					}
				} else {
					// Lấy phạm vi của MergeCell
					final var firstRowIdx = mergeCellRange.getFirstRow();
					final var lastRowIdx = mergeCellRange.getLastRow();

					var from = firstRowIdx; // Starting number
					final var to = lastRowIdx; // Ending number
					final List<String> listBacklogKeyExists = new ArrayList<>();
					while (from <= to) {
						final var curBacklogKey = getBacklogKeyVal(sheet.getRow(from));
						listBacklogKeyExists.add(curBacklogKey);
						from++;
					}
					final var newBacklogs = backlogs.stream().filter(
							x -> listBacklogKeyExists.stream().allMatch(k -> !StringUtils.equals(k, x.getKey())))
							.collect(Collectors.toList());

					final var curBacklogs = backlogs.stream()
							.filter(x -> listBacklogKeyExists.stream().anyMatch(k -> StringUtils.equals(k, x.getKey())))
							.collect(Collectors.toList());

					curRowCnt = listBacklogKeyExists.size();
					// Dịch chuyển các dòng
					final var startRowShift = curIdx + curRowCnt;
					numberOfRowsToShift = newBacklogs.size();
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

					// Điền thông tin cho record được thêm mới
					for (final BacklogDetail backlogDetail : newBacklogs) {

						final var curRowIdx = startRowShift + stepCnt;

						final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas, indexNo);

						allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

						stepCnt++;
					}

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
					var newMergedRegion = new CellRangeAddress(fRow, lRow, columnAIndex, columnAIndex);
					sheet.addMergedRegion(newMergedRegion);
					// Column "グループ Group"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnBIndex, columnBIndex);
					sheet.addMergedRegion(newMergedRegion);

					// Column "画面ID Screen ID"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnAnkenIndex, columnAnkenIndex);
					sheet.addMergedRegion(newMergedRegion);

					// Column "画面名 Screen Name"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex, columnScreenIndex);
					sheet.addMergedRegion(newMergedRegion);
					// Column "ステータス Status"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex, columnStatusIndex);
					sheet.addMergedRegion(newMergedRegion);
				}

				// tăng index xử lý sau khi xử lý thêm dòng
				final var increCnt = numberOfRowsToShift + curRowCnt - 1;
				i.addAndGet(increCnt);
			} else {
				// T/h new mới row
				// Thêm mới thông tin
				final var firstEntryOptional = groupedBacklogs.entrySet().stream().findFirst();

				if (!firstEntryOptional.isPresent()) {
					break;
				}
				final var firstEntry = firstEntryOptional.get();
				curBacklogParentKey = firstEntry.getKey(); // Lấy ra parent key
				final var backlogs = firstEntry.getValue();

				// Lấy ra ticket no
				final var curAnkenNo = backlogs.stream().findFirst().map(BacklogDetail::getAnkenNo).orElse("");
				final Integer totalRow = backlogs.size();
//					row.getCell(columnAIndex).setCellValue(curIdx + 1); // fill number no

				// Dịch chuyển các dòng
				final var startRowShift = curIdx + 1;
				numberOfRowsToShift = totalRow;
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

				if (totalRow != 1) {
					final var fRow = curIdx;
					final var lRow = fRow + numberOfRowsToShift - 1;
					// merge cell
					// Column No
					var newMergedRegion = new CellRangeAddress(fRow, lRow, columnAIndex, columnAIndex);
					sheet.addMergedRegion(newMergedRegion);
					// Column "グループ Group"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnBIndex, columnBIndex);
					sheet.addMergedRegion(newMergedRegion);
					setValForMergeCell(sheet, newMergedRegion, columnBIndex, curBacklogParentKey);

					// Column "画面ID Screen ID"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnAnkenIndex, columnAnkenIndex);
					sheet.addMergedRegion(newMergedRegion);
					setValForMergeCell(sheet, newMergedRegion, columnAnkenIndex, curAnkenNo);

					// Column "画面名 Screen Name"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex, columnScreenIndex);
					sheet.addMergedRegion(newMergedRegion);
					// Column "ステータス Status"
					newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex, columnStatusIndex);
					sheet.addMergedRegion(newMergedRegion);
				}
				// tăng index xử lý sau khi xử lý thêm dòng
				final var increCnt = numberOfRowsToShift - 1;
				i.addAndGet(increCnt);
			}
			// remove sau khi lay ra thong tin xu ly
			groupedBacklogs.remove(curBacklogParentKey);
			indexNo.incrementAndGet();

		}

		evaluate(workbook, sheet);

	}

	private void evaluate(final Workbook workbook, final Sheet sheet) {
		// Cập nhật lại công thức
		updatedTotalActualHoursFormula(sheet);

		updatedTotalFooterFormula(sheet);

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

		updatedTotalFooterFormula(sheet);

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
	 * @param projectCd
	 * @param datas
	 * @param workbook
	 */
	private void fillScheduleInfo(final String projectCd, final List<PjjyujiDetail> pds, final List<BacklogDetail> bds,
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
		case IFRONT -> pathIfrontTemplate;
		case SYMPHONIZER -> pathSymTemplate;
		default -> pathDefaultTemplate;
		};

		final var projectSchPath = Paths.get(String.format(projectScheduleTemplate, pathRootFolder, projectCd));
		if (!Files.exists(projectSchPath)) {
			Files.createDirectories(projectSchPath);
		}
		return projectSchPath;
	}

	public void createSchedule(final CustomerTarget projecType, final String projectCd, final List<PjjyujiDetail> pds,
			final List<BacklogDetail> bds) throws IOException {
		log.debug("Bat dau tao schedule: {}", projectCd);

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
				: BacklogExcel.class.getClassLoader().getResourceAsStream(scheduleTemplatePath);
				Workbook workbook = new XSSFWorkbook(fis)) {

			fillScheduleInfo(projectCd, pds, bds, workbook);

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
		final var fileName = StringUtils.replaceEach(templateFile, new String[] { "{projectCd}", "{range}" },
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
		return true;
	}

	private Map<String, List<BacklogDetail>> groupByParentKey(final List<BacklogDetail> details) {
		final Map<String, List<BacklogDetail>> groupedBacklogs = new HashMap<>();

		for (final BacklogDetail obj : details) {
			if (StringUtils.isNotBlank(obj.getParentKey())) {
				groupedBacklogs.computeIfAbsent(obj.getParentKey(), key -> new ArrayList<>()).add(obj);
			}
		}
		return groupedBacklogs;
	}

}

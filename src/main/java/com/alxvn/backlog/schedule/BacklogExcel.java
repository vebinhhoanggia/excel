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
import java.util.Comparator;
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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
	private static final int SCH_DEFAULT_ROW_CNT = 42;
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

	public int addTotalBacklogCol(final Sheet sheet, final FormulaEvaluator formulaEvaluator) {
		final var row = sheet.getRow(targetMonthRowIdx);

		final var columnIndexToInsert = columnStartDateInputIdx; // Position of the new column (0-based)
		final var shiftAmount = 1; // Number of columns to shift
		// Shift columns to the right
		sheet.shiftColumns(columnIndexToInsert, row.getLastCellNum() - 1, shiftAmount);

		for (final Row r : sheet) {
			for (final Cell c : r) {
				formulaEvaluator.evaluate(c);
			}
		}
		return shiftAmount;
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

	private void setAnkeNoValue(final Sheet sheet, final CellRangeAddress mergeCellRange, final String ankenNo) {
		final var firstRow = mergeCellRange.getFirstRow();
		final var fRow = firstRow;
		// Create a new cell within the merged region
		var r = sheet.getRow(fRow);
		if (r == null) {
			r = sheet.createRow(fRow);
		}
		var c = r.getCell(columnAnkenIndex);
		if (c == null) {
			c = r.createCell(columnAnkenIndex);
		}
		// Set the value for the cell
		c.setCellValue(ankenNo);
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

	private void setValForSingleCell(final Sheet sheet, final int rowIdx, final int colIdx, final String strVal) {
		final var fRow = rowIdx;
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

	private void setAnkeNoValue(final Sheet sheet, final int firstRow, final String ankenNo) {
		final var fRow = firstRow;
		// Create a new cell within the merged region
		var r = sheet.getRow(fRow);
		if (r == null) {
			r = sheet.createRow(fRow);
		}
		var c = r.getCell(columnAnkenIndex);
		if (c == null) {
			c = r.createCell(columnAnkenIndex);
		}
		// Set the value for the cell
		c.setCellValue(ankenNo);
	}

	private void insertNewRowInExistsCol(final Sheet sheet, final Row row, final CellRangeAddress mergeCellRange,
			final int numberOfRowsToShift, final String ankenNo) {
		if (mergeCellRange != null) {

			// Lấy phạm vi của MergeCell
			final var firstRow = mergeCellRange.getFirstRow();
			final var lastRow = mergeCellRange.getLastRow();

			final var rowIndex = mergeCellRange.getLastRow();

			// unmerge cell
			for (var i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
				final var mergedRegion = sheet.getMergedRegion(i);
				if (compareCellRangeAddresses(mergedRegion, mergeCellRange)) {
					sheet.removeMergedRegion(i);
				}
			}

			// Dịch chuyển các dòng
			sheet.shiftRows(rowIndex, sheet.getLastRowNum(), numberOfRowsToShift);

			final var fRow = firstRow;
			final var lRow = lastRow + numberOfRowsToShift;
			final var isHaveMergeCell = lRow - fRow > 1;
			if (isHaveMergeCell) {
				// merge lại cell
				// Column No
				var newMergedRegion = new CellRangeAddress(fRow, lRow, columnAIndex, columnAIndex);
				sheet.addMergedRegion(newMergedRegion);
				// Column Ticket
				newMergedRegion = new CellRangeAddress(fRow, lRow, columnAnkenIndex, columnAnkenIndex);
				sheet.addMergedRegion(newMergedRegion);
			}

			setAnkeNoValue(sheet, mergeCellRange, ankenNo);

			if (isHaveMergeCell) {
				// Column Screen
				var newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex, columnScreenIndex);
				sheet.addMergedRegion(newMergedRegion);
				// Column Status
				newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex, columnStatusIndex);
				sheet.addMergedRegion(newMergedRegion);
			}

			// Tạo dòng mới sau khi dịch chuyển
			for (var i = rowIndex; i <= rowIndex + numberOfRowsToShift - 1; i++) {
				final var newRow = sheet.createRow(i);
				cloneRowFormat(row, newRow);
			}
		}
	}

	private void insertNewRowBottom(final Sheet sheet, final Row bottomRow, final int cntRowInsert,
			final String ankenNo) {

		final var bottomRowIdx = bottomRow.getRowNum();

		final var rIdxStartShift = bottomRowIdx;
		// Dịch chuyển các dòng
		sheet.shiftRows(rIdxStartShift, sheet.getLastRowNum(), cntRowInsert);

		final var firstRow = rIdxStartShift;
		final var lastRow = firstRow + cntRowInsert - 1;

		if (cntRowInsert > 1) {
			// merge lại cell
			// Column No
			var newMergedRegion = new CellRangeAddress(firstRow, lastRow, columnAIndex, columnAIndex);
			sheet.addMergedRegion(newMergedRegion);
			// Column Ticket
			newMergedRegion = new CellRangeAddress(firstRow, lastRow, columnAnkenIndex, columnAnkenIndex);
			sheet.addMergedRegion(newMergedRegion);
			// Column Screen
			newMergedRegion = new CellRangeAddress(firstRow, lastRow, columnScreenIndex, columnScreenIndex);
			sheet.addMergedRegion(newMergedRegion);
			// Column Status
			newMergedRegion = new CellRangeAddress(firstRow, lastRow, columnStatusIndex, columnStatusIndex);
			sheet.addMergedRegion(newMergedRegion);
		}
		/**
		 * Copy format từ dòng trên cho các row được shift
		 */
		final var sourceRowFormat = sheet.getRow(bottomRowIdx - 1);
		// Tạo dòng mới sau khi dịch chuyển
		for (var i = bottomRowIdx; i <= bottomRowIdx + cntRowInsert - 1; i++) {
			final var newRow = sheet.createRow(i);
			cloneRowFormat(sourceRowFormat, newRow);
		}
		/**
		 * Thiết lập giá trị ankenNo vào merge cell
		 */
		// Create a new cell within the merged region
		var r = sheet.getRow(firstRow);
		if (r == null) {
			r = sheet.createRow(firstRow);
		}
		var c = r.getCell(columnAnkenIndex);
		if (c == null) {
			c = r.createCell(columnAnkenIndex);
		}
		// Set the value for the cell
		c.setCellValue(ankenNo);
	}

	public void insertNewRowForAll(final Sheet sheet, final int cntRowInsert) {

		Row bottomRow = null;
		for (final Row row : sheet) {
			if (isTotalRow(row)) {
				break;
			}
			bottomRow = row;
		}
		if (bottomRow == null) {
			return;
		}

		final var bottomRowIdx = bottomRow.getRowNum();

		final var rIdxStartShift = bottomRowIdx;
		if (cntRowInsert >= 1) {
			// Dịch chuyển các dòng
			sheet.shiftRows(rIdxStartShift, sheet.getLastRowNum(), cntRowInsert);
		}

		/**
		 * Copy format từ dòng trên cho các row được shift
		 */
		final var sourceRowFormat = sheet.getRow(bottomRowIdx - 1);
		// Tạo dòng mới sau khi dịch chuyển
		for (var i = bottomRowIdx; i <= bottomRowIdx + cntRowInsert - 1; i++) {
			final var newRow = sheet.createRow(i);
			cloneRowFormat(sourceRowFormat, newRow);
		}
	}

	private boolean isHeader(final Row row) {
		return row.getRowNum() <= targetDateRowIdx;
	}

	private boolean addRowForExistsAnkenVal(final String ankenNo, final Sheet sheet, final List<BacklogDetail> backlogs,
			final FormulaEvaluator formulaEvaluator) {
		var isExists = false;
		final var dataFormatter = new DataFormatter();
		final Integer totalRow = backlogs.size();
		for (final Row row : sheet) {
			// skip xử lý khi đang đọc các dòng header
			if (isHeader(row)) {
				continue;
			}
			final var cell = row.getCell(columnAnkenIndex);
			if (cell != null) {
				formulaEvaluator.evaluate(cell);
				final var formattedCellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);

				// Check if the value matches the desired value
				if (StringUtils.equals(ankenNo, formattedCellValue)) {
					final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(cell);
					if (mergeCellRange != null) {
						final var totalRowsOfCurrentTicket = mergeCellRange.getLastRow() - mergeCellRange.getFirstRow()
								+ 1;
						// Số lượng dòng cần dịch chuyển
						final var numberOfRowsToShift = totalRow - totalRowsOfCurrentTicket;
						if (numberOfRowsToShift > 0) {
							insertNewRowInExistsCol(sheet, row, mergeCellRange, numberOfRowsToShift, ankenNo);
						} else {
							setAnkeNoValue(sheet, mergeCellRange, ankenNo);
						}
						isExists = true;
						break;
					}
				}
			}
		}
		return isExists;
	}

	private boolean addRowForNewAnken(final String ankenNo, final Sheet sheet, final List<BacklogDetail> backlogs,
			final FormulaEvaluator formulaEvaluator) {
		var isExistsDefaultRow = false;
		final Integer totalRow = backlogs.size();
		final var dataFormatter = new DataFormatter();
		for (final Row row : sheet) {
			// skip xử lý khi đang đọc các dòng header
			if (isHeader(row)) {
				continue;
			}
			final var cell = row.getCell(columnAnkenIndex);
			if (cell != null) {
				formulaEvaluator.evaluate(cell);
				final var formattedCellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);

				// Check if the value matches the desired value
				if (StringUtils.isBlank(formattedCellValue)) {
					final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(cell);
					// t/h cell trong range
					if (mergeCellRange != null
							&& StringUtils.isBlank(ScheduleHelper.readContentFromMergedCells(sheet, mergeCellRange))) {
						final var totalRowsOfCurrentMergeCellBlank = mergeCellRange.getLastRow()
								- mergeCellRange.getFirstRow() + 1;

						// Số lượng dòng cần dịch chuyển
						final var numberOfRowsToShift = totalRow - totalRowsOfCurrentMergeCellBlank;
						if (numberOfRowsToShift > 0) {
							insertNewRowInExistsCol(sheet, row, mergeCellRange, numberOfRowsToShift, ankenNo);
						} else {
							setAnkeNoValue(sheet, mergeCellRange, ankenNo);
						}
						isExistsDefaultRow = true;
					}
					// t/h cell trong 1 row. Thưc hiện merge group lại
					if (mergeCellRange == null && StringUtils.isBlank(formattedCellValue)) {
						final var fRow = cell.getRowIndex();
						final var lRow = fRow + totalRow;
						// merge lại cell
						// Column No
						var newMergedRegion = new CellRangeAddress(fRow, lRow, columnAIndex, columnAIndex);
						sheet.addMergedRegion(newMergedRegion);
						// Column Ticket
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnAnkenIndex, columnAnkenIndex);
						sheet.addMergedRegion(newMergedRegion);

						setAnkeNoValue(sheet, newMergedRegion, ankenNo);

						// Column Screen
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex, columnScreenIndex);
						sheet.addMergedRegion(newMergedRegion);
						// Column Status
						newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex, columnStatusIndex);
						sheet.addMergedRegion(newMergedRegion);

						isExistsDefaultRow = true;
					}
				}
			}
		}
		return isExistsDefaultRow;
	}

	public void fillBacklogAndWrData(final Sheet sheet, final List<BacklogDetail> curBacklogs,
			final List<PjjyujiDetail> pds) {
		final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

//		final Map<String, List<BacklogDetail>> groupedBacklogs = curBacklogs.stream()
//				.collect(Collectors.groupingBy(BacklogDetail::getAnkenNo));

		final Map<String, List<BacklogDetail>> groupedBacklogs = curBacklogs.stream()
				.collect(Collectors.groupingBy(BacklogDetail::getAnkenNo,
						Collectors.collectingAndThen(Collectors.toList(),
								list -> list.stream()
										.sorted(Comparator.comparing(BacklogDetail::getActualStartDate,
												Comparator.nullsFirst(Comparator.naturalOrder())))
										.collect(Collectors.toList()))));

		final List<String> ankens = new ArrayList<>();
		for (final Map.Entry<String, List<BacklogDetail>> entry : groupedBacklogs.entrySet()) {
			final var ankenNo = entry.getKey();
			ankens.add(ankenNo);
		}
		final List<String> existedAnken = new ArrayList<>();
		final var i = new AtomicInteger(0);
		final var dataFormatter = new DataFormatter();
		final var sheetName = sheet.getSheetName();

		/**
		 * Thực hiện update record đã tồn tại ankenNo hoặc là thêm mới record cho anken
		 * mới
		 */
		for (final Row row : sheet) {
			final var curRowIdx = row.getRowNum();
			// skip xử lý khi đang đọc các dòng header
			if (isHeader(row)) {
				continue;
			}

			final var backlogKeyVal = StringUtils.trim(StringUtils.defaultString(dataFormatter.formatCellValue(
					row.getCell(CellReference.convertColStringToIndex(colBacklogIdChar)), formulaEvaluator)));

			final var cell = row.getCell(columnAnkenIndex);
			if (cell != null) {
				formulaEvaluator.evaluate(cell);
				final var formattedCellValue = StringUtils
						.trim(StringUtils.defaultString(dataFormatter.formatCellValue(cell, formulaEvaluator)));

				// Check if the value matches the desired value
				if (StringUtils.isBlank(formattedCellValue)) {
					final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(cell);
					if (mergeCellRange == null) { // thêm anken mới
						var curIdx = i.getAndIncrement();
						final var ankenNo = ankens.get(curIdx);
						var curAnkenNo = ankenNo;
						// Skip if exists in existedAnken;
						if (existedAnken.stream().anyMatch(x -> StringUtils.equals(x, ankenNo))) {
							curIdx = i.getAndIncrement();
							curAnkenNo = ankens.get(curIdx);
						}

						final var backlogs = groupedBacklogs.get(curAnkenNo);
						final Integer totalRow = backlogs.size();
						row.getCell(columnAIndex).setCellValue(curIdx + 1); // fill number no
						if (totalRow == 1) {
							final var rowStart = row.getRowNum();
							setAnkeNoValue(sheet, rowStart, curAnkenNo);
							fillDataAfterMergeCell(sheet, curAnkenNo, rowStart, backlogs, pds);
						} else {

							final var fRow = cell.getRowIndex();
							final var lRow = fRow + totalRow - 1;
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

							setAnkeNoValue(sheet, newMergedRegion, curAnkenNo);

							// Column "画面名 Screen Name"
							newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex, columnScreenIndex);
							sheet.addMergedRegion(newMergedRegion);
							// Column "ステータス Status"
							newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex, columnStatusIndex);
							sheet.addMergedRegion(newMergedRegion);
							final var rowStart = newMergedRegion.getFirstRow();
							fillDataAfterMergeCell(sheet, curAnkenNo, rowStart, backlogs, pds);
						}
					} else {
						final var curAnkenNo = ScheduleHelper.readContentFromMergedCells(sheet, mergeCellRange);
						if (StringUtils.isNotBlank(curAnkenNo)) {
							// update anken có sẵn mà có trong merge cell
							// t/h merge cell và có giá trị anken no
							final var rowStart = mergeCellRange.getLastRow();
							if (curRowIdx >= rowStart) {
								i.getAndIncrement();
							}
							if (existedAnken.stream().allMatch(x -> !StringUtils.equals(x, curAnkenNo))) {
								existedAnken.add(curAnkenNo);
							}
							final var backlogs = groupedBacklogs.get(curAnkenNo);
							final var backlog = backlogs.stream()
									.filter(x -> StringUtils.equals(x.getKey(), backlogKeyVal)).findFirst()
									.orElse(null);
							final var wrTargets = pds.stream().filter(x -> filterPredicate(x, sheetName, curAnkenNo))
									.collect(Collectors.toList());
							final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlog, wrTargets);
							pds.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule
						} else {
							// khong doi ung case nay. Bat buoc merge cell phai co anken no
						}
					}
				} else {
					// update anken có sẵn mà không có merge cell
					final var curAnkenNo = formattedCellValue;
					final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(cell);
					if (mergeCellRange == null) { // t/h 1 record
						i.getAndIncrement();
					} else { // t/h nhieu record
						final var rowStart = mergeCellRange.getLastRow();
						if (curRowIdx >= rowStart) {
							i.getAndIncrement();
						}
					}
					if (existedAnken.stream().allMatch(x -> !StringUtils.equals(x, curAnkenNo))) {
						existedAnken.add(curAnkenNo);
					}
					final var backlogs = groupedBacklogs.get(curAnkenNo);
					final var backlog = backlogs.stream().filter(x -> StringUtils.equals(x.getKey(), backlogKeyVal))
							.findFirst().orElse(null);
					final var wrTargets = pds.stream().filter(x -> filterPredicate(x, sheetName, curAnkenNo))
							.collect(Collectors.toList());
					final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlog, wrTargets);
					pds.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule
				}
			}
			if (i.get() >= ankens.size()) {
				break; // Exit the while loop
			}
		}
		final List<BacklogDetail> newDataBacklogs = new ArrayList<>();

		/**
		 * Thực hiện loop từng anken để insert thêm dòng mới và điền thông tin.
		 *
		 */
		for (final String currentAnkenNo : existedAnken) {
			// Them record moi vao cac anken da ton tai trong schedule
			for (final Row row : sheet) {
				final var curRowIdx = row.getRowNum();
				// skip xử lý khi đang đọc các dòng header
				if (isHeader(row)) {
					continue;
				}

				final var cell = row.getCell(columnAnkenIndex);
				if (cell != null) {
					formulaEvaluator.evaluate(cell);
					final var formattedCellValue = StringUtils
							.trim(StringUtils.defaultString(dataFormatter.formatCellValue(cell, formulaEvaluator)));
					final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(cell);
					if (StringUtils.isBlank(formattedCellValue) && mergeCellRange == null) {
						// Đã đi đến record cuối. Thoát vòng lặp.
						break;
					}
					// Tìm kiếm record sẽ thực hiện thêm
					final var ismatchRow = StringUtils.equals(currentAnkenNo, formattedCellValue)
							|| mergeCellRange != null && StringUtils.equals(currentAnkenNo,
									ScheduleHelper.readContentFromMergedCells(sheet, mergeCellRange));
					if (ismatchRow) {
						final var backlogKeyVal = StringUtils.trim(StringUtils.defaultString(dataFormatter
								.formatCellValue(row.getCell(CellReference.convertColStringToIndex(colBacklogIdChar)),
										formulaEvaluator)));
						final var curAnkenNo = Objects.toString(formattedCellValue,
								ScheduleHelper.readContentFromMergedCells(sheet, mergeCellRange));

						var totalRow = 0;

						// T/h cell trong range
						if (mergeCellRange != null) {
							final var backlogs = groupedBacklogs.get(curAnkenNo);

							final var firstRowIdx = mergeCellRange.getFirstRow();
							final var lastRowIdx = mergeCellRange.getLastRow();

							var from = firstRowIdx; // Starting number
							final var to = lastRowIdx; // Ending number
							final List<String> listBacklogKeyExists = new ArrayList<>();
							while (from <= to) {
								final var curBacklogKey = StringUtils
										.trim(StringUtils
												.defaultString(dataFormatter.formatCellValue(
														sheet.getRow(from)
																.getCell(CellReference
																		.convertColStringToIndex(colBacklogIdChar)),
														formulaEvaluator)));
								listBacklogKeyExists.add(curBacklogKey);
								from++;
							}
							final var newBacklogs = backlogs.stream()
									.filter(x -> listBacklogKeyExists.stream()
											.allMatch(k -> !StringUtils.equals(k, x.getKey())))
									.collect(Collectors.toList());

							totalRow = newBacklogs.size();

							if (totalRow > 0) {
								// luu lai danh sach de update data
								newBacklogs.addAll(CollectionUtils.emptyIfNull(newBacklogs).stream().toList());
								final var mergeTotalRow = mergeCellRange.getLastRow() - mergeCellRange.getFirstRow()
										+ 1;
								// Số lượng dòng cần dịch chuyển
								final var numberOfRowsToShift = totalRow - mergeTotalRow;
								if (numberOfRowsToShift > 0) {
									insertNewRowInExistsCol(sheet, row, mergeCellRange, numberOfRowsToShift,
											curAnkenNo);
								} else {
									// Không phát sinh thêm record
								}
							}
						}
						// t/h cell trong 1 row. Thưc hiện them dong, merge group lại
						if (mergeCellRange == null) {
							final var backlogs = groupedBacklogs.get(curAnkenNo);
							final var newBacklogs = CollectionUtils.emptyIfNull(backlogs).stream()
									.filter(x -> !StringUtils.equals(x.getKey(), backlogKeyVal))
									.collect(Collectors.toList());
							totalRow = newBacklogs.size();
							if (totalRow > 0) {
								final var numberOfRowsToShift = totalRow;

								sheet.shiftRows(curRowIdx + 1, sheet.getLastRowNum(), numberOfRowsToShift);
								final var fRow = curRowIdx;
								final var lRow = fRow + totalRow;
								// merge lại cell
//								// Column No
								var newMergedRegion = new CellRangeAddress(fRow, lRow, columnAIndex, columnAIndex);
								sheet.addMergedRegion(newMergedRegion);
								// Column Ticket
								newMergedRegion = new CellRangeAddress(fRow, lRow, columnAnkenIndex, columnAnkenIndex);
								sheet.addMergedRegion(newMergedRegion);

								setAnkeNoValue(sheet, newMergedRegion, curAnkenNo);

								// Column Screen
								newMergedRegion = new CellRangeAddress(fRow, lRow, columnScreenIndex,
										columnScreenIndex);
								sheet.addMergedRegion(newMergedRegion);
								// Column Status
								newMergedRegion = new CellRangeAddress(fRow, lRow, columnStatusIndex,
										columnStatusIndex);
								sheet.addMergedRegion(newMergedRegion);

//								 Tạo dòng mới sau khi dịch chuyển
								for (var k = curRowIdx + 1; k <= curRowIdx + numberOfRowsToShift; k++) {
									final var newRow = sheet.createRow(k);
									cloneRowFormat(row, newRow);
								}
							}
						}
						// Sau khi thêm thì thoát vòng lặp. Thực hiện tiếp anken mới.
						break;
					}
				}
			}
		}
	}

	@SuppressWarnings("unused")
	private void crateRowBacklog(final String ankenNo, final Sheet sheet, final List<BacklogDetail> backlogs,
			final FormulaEvaluator formulaEvaluator) {
		final var isExists = addRowForExistsAnkenVal(ankenNo, sheet, backlogs, formulaEvaluator);
		if (isExists) {
			// Cập nhật thông tin vào các row của ticket/anken
			fillBacklogData(ankenNo, sheet, backlogs, formulaEvaluator);
			return;
		}
		/*
		 * Trường hợp chưa tồn tại row thì tìm group default sau đó thêm cho đủ row. T/h
		 * không có group default thì nếu có đủ row thì thực hiện tạo group.
		 */
		final var isExistsDefaultRow = addRowForNewAnken(ankenNo, sheet, backlogs, formulaEvaluator);
		if (isExistsDefaultRow) {
			// Cập nhật thông tin vào các row của ticket/anken
			fillBacklogData(ankenNo, sheet, backlogs, formulaEvaluator);
			return;
		}
		/*
		 * T/h không có default row thì tạo mới
		 */
		final Integer totalRow = backlogs.size();
		Row bottomRow = null;
		for (final Row row : sheet) {
			if (isTotalRow(row)) {
				break;
			}
			bottomRow = row;
		}
		if (bottomRow != null) {
			insertNewRowBottom(sheet, bottomRow, totalRow, ankenNo);
		}
		// Cập nhật thông tin vào các row của ticket/anken
		fillBacklogData(ankenNo, sheet, backlogs, formulaEvaluator);
	}

	public Date toDate(final LocalDate localDate) {
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

	private boolean filterPredicate(final PjjyujiDetail pjjyujiDetail, final String sheetName, final String ankenNo) {
//		final var progressName = pjjyujiDetail.getProcess().getName();
		final var progressCd = pjjyujiDetail.getProcess().getCode();
		if (StringUtils.equals(sheetName, "pg_spec")) {
			return StringUtils.equals(ankenNo, pjjyujiDetail.getAnkenNo())
					&& StringUtils.equals(progressCd, WorkingPhase.ID45.getCode());
		}
		if (StringUtils.equals(sheetName, "pg_bug")) {
			return StringUtils.equals(ankenNo, pjjyujiDetail.getAnkenNo())
					&& StringUtils.equals(progressCd, WorkingPhase.ID43.getCode());
		}
		return StringUtils.equals(ankenNo, pjjyujiDetail.getAnkenNo());
	}

	private Cell getCell(final Row curRow, final String colChar) {
		final var cellIdx = CellReference.convertColStringToIndex(colChar);
		var curCel = curRow.getCell(cellIdx);
		if (curCel == null) {
			curCel = curRow.createCell(cellIdx);
		}
		return curCel;
	}

	private void fillBacklogDataForRow(final Row curRow, final BacklogDetail backlogDetail) {
		// "グループ Group"
		final var parentKey = Optional.ofNullable(backlogDetail).map(BacklogDetail::getParentKey)
				.orElse(StringUtils.EMPTY);
		var curCel = getCell(curRow, columnBCharacter);
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

	private void fillDataAfterMergeCell(final Sheet sheet, final String ankenNo, final int rowStart,
			final Collection<BacklogDetail> backlogs, final List<PjjyujiDetail> pds) {

		var curIdx = 0;
		final var sheetName = sheet.getSheetName();
		final var wrTargets = pds.stream().filter(x -> filterPredicate(x, sheetName, ankenNo))
				.collect(Collectors.toList());
		for (final BacklogDetail backlogDetail : backlogs) {

			final var curRowIdx = rowStart + curIdx;

			final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, wrTargets);

			wrTargets.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule
			pds.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

			curIdx++;
		}
	}

	private Collection<PjjyujiDetail> fillDataForRow(final Sheet sheet, final int curRowIdx,
			final BacklogDetail backlogDetail, final List<PjjyujiDetail> wrTargets) {
		if (backlogDetail == null) {
			return Collections.emptyList();
		}
		final var curRow = sheet.getRow(curRowIdx);

		fillBacklogDataForRow(curRow, backlogDetail);

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

	private void fillBacklogData(final String ankenNo, final Sheet sheet, final List<BacklogDetail> backlogs,
			final FormulaEvaluator formulaEvaluator) {
		final var dataFormatter = new DataFormatter();
		for (final Row row : sheet) {
			// skip xử lý khi đang đọc các dòng header
			if (isHeader(row)) {
				continue;
			}
			final var cell = row.getCell(columnAnkenIndex);
			if (cell != null) {
				formulaEvaluator.evaluate(cell);
				final var formattedCellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);

				// Check if the value matches the desired value
				if (StringUtils.equals(ankenNo, formattedCellValue)) {
					final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(cell);
					var rowStart = row.getRowNum();
					if (mergeCellRange != null) {
						rowStart = mergeCellRange.getFirstRow();
					}
					var curIdx = 0;
					for (final BacklogDetail backlogDetail : backlogs) {
						final var curRowIdx = rowStart + curIdx;
						final var curRow = sheet.getRow(curRowIdx);
						// 工程 Operation
						var curCel = curRow.getCell(CellReference.convertColStringToIndex(colOperationChar));
						curCel.setCellValue(
								getOperation(backlogDetail).map(WorkingPhase::getName).orElse(StringUtils.EMPTY));
						// 担当 PIC
						curCel = curRow.getCell(CellReference.convertColStringToIndex(columnPicCharacter));
						curCel.setCellValue(backlogDetail.getMailId());
						// ステータス Status
						// curCel =
						// curRow.getCell(CellReference.convertColStringToIndex(columnStatusCharacter));
						// curCel.setCellValue(backlogDetail.getStatus());

						// "予定 Schedule"
						// 工数 Hours
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colExpectHousrChar));
						curCel.setCellValue(Optional.ofNullable(backlogDetail.getEstimatedHours())
								.orElse(BigDecimal.ZERO).doubleValue());
						// 開始日 Begin
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colExpectStartYmdChar));
						curCel.setCellValue(toDate(backlogDetail.getExpectedStartDate()));
						// 完了日 End
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colExpectEndYmdChar));
						curCel.setCellValue(toDate(backlogDetail.getExpectedDueDate()));
						// 納品日 Delivery
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colExpectDeliveryYmdChar));
						curCel.setCellValue(toDate(backlogDetail.getExpectedDeliveryDate()));

						// "実績 Actual"
						// 工数 Hours
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colActualTotalHoursBacklogChar));
						curCel.setCellValue(Optional.ofNullable(backlogDetail.getActualHours()).orElse(BigDecimal.ZERO)
								.doubleValue());
						// 開始日 Begin
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colActStartYmdChar));
						curCel.setCellValue(toDate(backlogDetail.getActualStartDate()));
						// 完了日 End
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colActEndYmdChar));
						curCel.setCellValue(toDate(backlogDetail.getActualDueDate()));
						// 進捗 Progress
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colActProgressChar));
						curCel.setCellValue(getProgress(backlogDetail.getProgress()));
						// 納品日 Delivery
						curCel = curRow.getCell(CellReference.convertColStringToIndex(colActDeliveryYmdChar));
						curCel.setCellValue(toDate(backlogDetail.getActualDeliveryDate()));

						curIdx++;
					}
					break;
				}
			}
		}
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

	public void reUpdateFormatCondition(final Sheet sheet, final int numOfAddRow, final int numOfAddCol) {

		final var formatting = sheet.getSheetConditionalFormatting();

		formatting.getNumConditionalFormattings();

		// Get the existing ConditionalFormattingRule
		final var cf = formatting.getConditionalFormattingAt(0);
		final var existingRule = cf.getRule(0);

		// Update the range to new range
		final var oldRange = cf.getFormattingRanges()[0];
		final var newRange = new CellRangeAddress(oldRange.getFirstRow(), oldRange.getLastRow() + numOfAddRow,
				oldRange.getFirstColumn(), oldRange.getLastColumn() + numOfAddCol);

		// Remove the existing formatting for the old range
		formatting.removeConditionalFormatting(0);

		// Apply the existing ConditionalFormattingRule to the new range
		formatting.addConditionalFormatting(new CellRangeAddress[] { newRange }, existingRule);

		// ... your code to process the number of rules
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

	private void fillDataForSheet(final Workbook workbook, final Sheet sheet, final List<BacklogDetail> allBacklogs,
			final List<PjjyujiDetail> allWrDatas) {
		if (CollectionUtils.isEmpty(allBacklogs)) {
			return;
		}
		final var groupedBacklogs = groupByParentKey(allBacklogs);
		final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		final var dataFormatter = new DataFormatter();
		final var i = new AtomicInteger(0);
		while (MapUtils.isNotEmpty(groupedBacklogs)) {
			sheet.getSheetName();
			final var curIdx = i.getAndIncrement();
			final var row = sheet.getRow(curIdx);
			// skip xử lý khi đang đọc các dòng header
			if (row == null || isHeader(row)) {
				continue;
			}
			final var groupCell = row.getCell(columnBIndex);
			if (groupCell != null) {
				formulaEvaluator.evaluate(groupCell);
				var curBacklogParentKey = dataFormatter.formatCellValue(groupCell, formulaEvaluator);
				final var isExists = StringUtils.isNotBlank(curBacklogParentKey)
						&& groupedBacklogs.containsKey(curBacklogParentKey);

				final var mergeCellRange = ScheduleHelper.getMergedRegionForCell(groupCell);
				final var isSingleRecord = mergeCellRange == null;
				var numberOfRowsToShift = 0;
				var curRowCnt = 0;
				// T/h tồn tại thực hiện cập nhật thông tin, thêm dòng mới, merge cell lại
				if (isExists) {
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
							sheet.shiftRows(startRowShift, sheet.getLastRowNum(), numberOfRowsToShift);

							// Tạo dòng mới sau khi dịch chuyển
							for (var j = startRowShift; j < startRowShift + numberOfRowsToShift; j++) {
								final var newRow = sheet.createRow(j);
								cloneRowFormat(row, newRow);
							}
						}

						// Cập nhật thông tin cho record đã tồn tại
						if (curBacklog != null) {
							final var wrRemoveEles = fillDataForRow(sheet, curIdx, curBacklog, allWrDatas);

							allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule
						}

						// Điền thông tin cho record được thêm mới
						var stepCnt = 0;
						for (final BacklogDetail backlogDetail : newBacklogs) {

							final var curRowIdx = startRowShift + stepCnt;

							final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas);

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

						final var curBacklogs = backlogs.stream().filter(
								x -> listBacklogKeyExists.stream().anyMatch(k -> StringUtils.equals(k, x.getKey())))
								.collect(Collectors.toList());

						curRowCnt = listBacklogKeyExists.size();
						// Dịch chuyển các dòng
						final var startRowShift = curIdx + curRowCnt;
						numberOfRowsToShift = newBacklogs.size();
						if (numberOfRowsToShift > 0) {
							sheet.shiftRows(startRowShift, sheet.getLastRowNum(), numberOfRowsToShift);

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

							final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas);

							allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

							stepCnt++;
						}

						// Điền thông tin cho record được thêm mới
						for (final BacklogDetail backlogDetail : newBacklogs) {

							final var curRowIdx = startRowShift + stepCnt;

							final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas);

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
				} else {
					curRowCnt = 1;

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
					final var startRowShift = curIdx + curRowCnt;
					numberOfRowsToShift = totalRow - curRowCnt;
					if (numberOfRowsToShift > 0) {
						sheet.shiftRows(startRowShift, sheet.getLastRowNum(), numberOfRowsToShift);

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

						final var wrRemoveEles = fillDataForRow(sheet, curRowIdx, backlogDetail, allWrDatas);

						allWrDatas.removeAll(wrRemoveEles); // remove các record đã ghi vào schedule

						stepCnt++;
					}

					if (totalRow != 1) {
						final var fRow = curIdx;
						final var lRow = fRow + numberOfRowsToShift;
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

				}
				// remove sau khi lay ra thong tin xu ly
				groupedBacklogs.remove(curBacklogParentKey);

				// tăng index xử lý sau khi xử lý thêm dòng
				i.addAndGet(numberOfRowsToShift + curRowCnt - 1);
			}
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

	public Path getLastSchedule(final Path projectSchPath) throws IOException {
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

	public List<String> getCurrentBacklogKeyExists(final Sheet sheet) {
		final var formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var dataFormatter = new DataFormatter();

		final List<String> result = new ArrayList<>();
		for (final Row row : sheet) {
			if (isHeader(row)) {
				continue;
			}
			final var backlogKeyVal = StringUtils.trim(StringUtils.defaultString(dataFormatter.formatCellValue(
					row.getCell(CellReference.convertColStringToIndex(colBacklogIdChar)), formulaEvaluator)));
			if (StringUtils.isNotBlank(backlogKeyVal)) {
				result.add(backlogKeyVal);
			}
			if (isTotalRow(row)) {
				break;
			}
		}
		return result;
	}
}

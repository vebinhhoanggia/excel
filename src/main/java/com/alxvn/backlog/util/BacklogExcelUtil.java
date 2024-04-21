/**
 *
 */
package com.alxvn.backlog.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
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

import com.alxvn.backlog.BacklogService;
import com.alxvn.backlog.dto.BacklogDetail;
import com.alxvn.backlog.dto.PjjyujiDetail;

/**
 *
 */
public class BacklogExcelUtil {
	private static final String columnACharacter = "A";
	private static final String columnAnkenCharacter = "C";
	private static final String columnScreenCharacter = "D";
	private static final String colTOperationChar = "E";
	private static final String columnPicCharacter = "F";
	private static final String columnTotalCharacter = "F";
	private static final String columnStatusCharacter = "G";
	private static final String colExpectHousrChar = "H";
	private static final String colExpectStartYmdChar = "I";
	private static final String colExpectEndYmdChar = "J";
	private static final String colExpectDeliveryYmdChar = "K";
	private static final String colActualTotalHoursBacklogChar = "L";
	private static final String colActHousrChar = "M";
	private static final String colActStartYmdChar = "N";
	private static final String colActEndYmdChar = "O";
	private static final String colActProgressChar = "P";
	private static final String colActDeliveryYmdChar = "Q";

	private static final String totalCharacter = "Total";
	private static final int columnAIndex = CellReference.convertColStringToIndex(columnACharacter);
	private static final int columnAnkenIndex = CellReference.convertColStringToIndex(columnAnkenCharacter);
	private static final int columnScreenIndex = CellReference.convertColStringToIndex(columnScreenCharacter);
	private static final int columnStatusIndex = CellReference.convertColStringToIndex(columnStatusCharacter);
	private static final int targetMonthRowIdx = 7;
	private static final int targetDateRowIdx = 8;
	private static final int columnStartDateInputIdx = 17;
	private static final DateTimeFormatter FORMATTER_YYYYMMDD = DateTimeFormatter.ofPattern("yyyy/MM/dd");

	private static final String templateFile = "QDA-0222a_プロジェクト管理表_{projectCd}.xlsm";
	private static final String templateTotalActHours = "SUM(R{rIdx}:{cName}{rIdx})";
	private static final String templateNextDateFormula = "{preCol}+1";

	private static final String templateTotalHours = "SUM({cName}{rIdxS}:{cName}{rIdxE})";
	private static final String templateTotalHoursForPic = "SUMIF($G${rIdxS}:$G${rIdxE},$G{rIdxTarget},{cName}${rIdxS}:{cName}${rIdxE})";

	private static final String scheduleTemplatePath = "templates/QDA-0222a_プロジェクト管理表.xlsm";

	private static final String issuTypeSpec = "課題(委託)";
	private static final String issuTypeBug = "バグ";

	private boolean compareCellRangeAddresses(CellRangeAddress range1, CellRangeAddress range2) {
		// Compare the first row, last row, first column, and last column
		return range1.getFirstRow() == range2.getFirstRow() && range1.getLastRow() == range2.getLastRow();
	}

	private boolean isTotalRow(Row row) {
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
	private void updatedTotalActualHoursFormula(final Sheet sheet, final FormulaEvaluator formulaEvaluator) {

		final var columnIndex = CellReference.convertColStringToIndex(colActHousrChar);
		for (final Row row : sheet) {
			final var chr = ScheduleHelper.convertColumnIndexToName(row.getLastCellNum());
			final var rNum = row.getRowNum();
			if (rNum >= 9) {
				final var cell = row.getCell(columnIndex);
				if (cell != null) {
					final var adjustedFormula = StringUtils.replaceEach(templateTotalActHours,
							new String[] { "{rIdx}", "{cName}" }, new String[] { String.valueOf(rNum + 1), chr });
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
	private void updatedTotalFooterFormula(final Sheet sheet, final FormulaEvaluator formulaEvaluator) {
		// TODO
	}

	public String getCellValue(Cell cell) {
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

	private String evaluateFormulaCell(Cell cell) {
		final var formulaEvaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
		final var cellValue = formulaEvaluator.evaluate(cell);
		return getCellValueFromFormulaResult(cellValue);
	}

	private String getCellValueFromFormulaResult(CellValue cellValue) {
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
	private void cloneRowFormat(Row sourceRow, Row newRow, FormulaEvaluator formulaEvaluator) {
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

	private void setAnkeNoValue(Sheet sheet, final CellRangeAddress mergeCellRange, String ankenNo) {
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

	private void insertNewRowInExistsCol(Sheet sheet, Row row, final CellRangeAddress mergeCellRange,
			FormulaEvaluator formulaEvaluator, int numberOfRowsToShift, final String ankenNo) {
		if (mergeCellRange != null) {
			System.out.println("Cell found at row " + (row.getRowNum() + 1) + ", column " + columnAnkenCharacter
					+ ", range: " + mergeCellRange.formatAsString());

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
				cloneRowFormat(row, newRow, formulaEvaluator);
			}
			System.out.println("New row with formulas created successfully.");
		}
	}

	private void insertNewRowBottom(Sheet sheet, Row bottomRow, FormulaEvaluator formulaEvaluator, int cntRowInsert,
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
			cloneRowFormat(sourceRowFormat, newRow, formulaEvaluator);
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
		System.out.println("New row Bottom with formulas created successfully.");
	}

	private boolean isHeader(Row row) {
		return row.getRowNum() < targetDateRowIdx;
	}

	private void addRowForInsertData(String ankenNo, Sheet sheet, List<BacklogDetail> backlogs,
			FormulaEvaluator formulaEvaluator) {
		final var dataFormatter = new DataFormatter();
		var isExists = false;
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
						System.out.println("Cell found at row " + (row.getRowNum() + 1) + ", column "
								+ columnAnkenCharacter + ", range: " + mergeCellRange.formatAsString());
						final var totalRowsOfCurrentTicket = mergeCellRange.getLastRow() - mergeCellRange.getFirstRow()
								+ 1;
						final var numberOfRowsToShift = totalRow - totalRowsOfCurrentTicket; // Số lượng dòng cần dịch
						// chuyển
						if (numberOfRowsToShift > 0) {
							insertNewRowInExistsCol(sheet, row, mergeCellRange, formulaEvaluator, numberOfRowsToShift,
									ankenNo);
						} else {
							setAnkeNoValue(sheet, mergeCellRange, ankenNo);
						}
						isExists = true;
						break;
					}
				}
			}
		}
		if (isExists) {
			// Cập nhật thông tin vào các row của ticket/anken
			fillDataForRow(ankenNo, sheet, backlogs, formulaEvaluator);
			return;
		}
		/*
		 * Trường hợp chưa tồn tại row thì tìm group default sau đó thêm cho đủ row
		 */
		var isExistsDefaultRow = false;
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
					if (mergeCellRange != null
							&& StringUtils.isBlank(ScheduleHelper.readContentFromMergedCells(sheet, mergeCellRange))) {
						final var totalRowsOfCurrentMergeCellBlank = mergeCellRange.getLastRow()
								- mergeCellRange.getFirstRow() + 1;

						// Số lượng dòng cần dịch chuyển
						final var numberOfRowsToShift = totalRow - totalRowsOfCurrentMergeCellBlank;
						if (numberOfRowsToShift > 0) {
							insertNewRowInExistsCol(sheet, row, mergeCellRange, formulaEvaluator, numberOfRowsToShift,
									ankenNo);
						} else {
							setAnkeNoValue(sheet, mergeCellRange, ankenNo);
						}
						isExistsDefaultRow = true;
						break;
					}
				}
			}
		}
		if (isExistsDefaultRow) {
			// Cập nhật thông tin vào các row của ticket/anken
			fillDataForRow(ankenNo, sheet, backlogs, formulaEvaluator);
			return;
		}
		/*
		 * T/h không có default row thì tạo mới
		 */
		Row bottomRow = null;
		for (final Row row : sheet) {
			if (isTotalRow(row)) {
				break;
			}
			bottomRow = row;
		}
		if (bottomRow != null) {
			insertNewRowBottom(sheet, bottomRow, formulaEvaluator, totalRow, ankenNo);
		}
		// Cập nhật thông tin vào các row của ticket/anken
		fillDataForRow(ankenNo, sheet, backlogs, formulaEvaluator);
	}

	public Date toDate(LocalDate localDate) {
		if (localDate == null) {
			return null;
		}
		// Convert LocalDate to Date
		var localDateTime = localDate.atStartOfDay();
		var zonedDateTime = localDateTime.atZone(ZoneId.systemDefault());
		var date = Date.from(zonedDateTime.toInstant());
		return date;
	}

	private String getOperation(String backlogIssueType) {
		if (StringUtils.isBlank(backlogIssueType)) {
			return StringUtils.EMPTY;
		}
		return backlogIssueType;
	}

	private double getProgress(String backlogProgress) {
		if (StringUtils.isBlank(backlogProgress)) {
			return 0;
		}
		double value = NumberUtils.toInt(backlogProgress);
		var percent = 0.01;

		return value * percent;
	}

	private void fillDataForRow(String ankenNo, Sheet sheet, List<BacklogDetail> backlogs,
			FormulaEvaluator formulaEvaluator) {
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
					if (mergeCellRange != null) {
						var rowStart = mergeCellRange.getFirstRow();
						var curIdx = 0;
						for (BacklogDetail backlogDetail : backlogs) {
							var curRowIdx = rowStart + curIdx;
							var curRow = sheet.getRow(curRowIdx);
							// 工程 Operation
							var curCel = curRow.getCell(CellReference.convertColStringToIndex(colTOperationChar));
							curCel.setCellValue(getOperation(backlogDetail.getIssueType()));
							// 担当 PIC
							curCel = curRow.getCell(CellReference.convertColStringToIndex(columnPicCharacter));
							curCel.setCellValue(backlogDetail.getMailId());
							// ステータス Status
//							curCel = curRow.getCell(CellReference.convertColStringToIndex(columnStatusCharacter));
//							curCel.setCellValue(backlogDetail.getStatus());

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
							curCel = curRow
									.getCell(CellReference.convertColStringToIndex(colActualTotalHoursBacklogChar));
							curCel.setCellValue(Optional.ofNullable(backlogDetail.getActualHours())
									.orElse(BigDecimal.ZERO).doubleValue());
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
	}

	private void standardizedRangeInput(Sheet sheet, FormulaEvaluator formulaEvaluator, YearMonth targetYmS,
			YearMonth targetYmE) {
		final var row = sheet.getRow(targetMonthRowIdx);
		String lastTarget = null;
		for (final Cell cell : row) {
			if (cell == null || cell.getColumnIndex() < columnStartDateInputIdx) {
				continue;
			}
			final var cellVal = StringUtils.trim(ScheduleHelper.readContentCell(sheet, cell));
			if (StringUtils.isNotBlank(cellVal)) {
				lastTarget = cellVal;
				System.out.println("LastTarget: " + lastTarget);
			}
		}
		if (StringUtils.isBlank(lastTarget)) { // Check sheet is from template
			var currentYm = targetYmS;
			Boolean isFirst = true;
			while (currentYm.isBefore(targetYmE.plusMonths(1))) {
				addColInput(sheet, formulaEvaluator, currentYm, isFirst);
				currentYm = currentYm.plusMonths(1);
				isFirst = false;
			}
		}
	}

	private void addColInput(final Sheet sheet, final FormulaEvaluator formulaEvaluator, YearMonth targetYm,
			final boolean isFirstTargetMonth) {
		final var row = sheet.getRow(targetMonthRowIdx);
		final var lastColumnIndex = row.getLastCellNum() - 1;

		// Check if there are any rows in the sheet
		if (sheet.getLastRowNum() < 0) {
			System.out.println("Sheet is empty.");
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
		System.out.println("Add column for new target successfully.");
	}

	public void reUpdateFormatCondition(Sheet sheet, int numOfAddRow, int numOfAddCol) {

		final var formatting = sheet.getSheetConditionalFormatting();

		// Get the number of conditional formatting rules
		final var numFormattingRules = formatting.getNumConditionalFormattings();
		System.out.println("Number of conditional formatting rules: " + numFormattingRules);

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
	public static void main(String[] args) {
		var obj = new BacklogExcelUtil();
		final var wrPath = "templates/pjjyuji_data_csv_20240415.csv";
		final var backlogPath = "templates/Backlog-Issues-20240415-1157.csv";
		obj.createScheduleFromBacklog(wrPath, backlogPath);
	}

	public void createScheduleFromBacklog(String wrPath, String backlogPath) {
		final var backlogService = new BacklogService();
		try {
			backlogService.stastics(wrPath, backlogPath);
		} catch (final Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void doSample() {
		final var ankenNoUpdate = "#29033";
		final var ankenNoInsert = "SYMPHO-001";
		final var now = YearMonth.now();
		List<BacklogDetail> l = new ArrayList<>(6);
		try (var fis = BacklogExcelUtil.class.getClassLoader().getResourceAsStream(scheduleTemplatePath);
				Workbook workbook = new XSSFWorkbook(fis)) {
			final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
			final var sheetIterator = workbook.sheetIterator();
			while (sheetIterator.hasNext()) {
				final var sheet = sheetIterator.next();
				// Kiểm tra là sheet điền schedule
				if (!ScheduleHelper.isScheduleSheet(sheet)) {
					continue;
				}
				System.out.println("Sheet: " + sheet.getSheetName());

//				addTotalBacklogCol(sheet, formulaEvaluator);

				// Thêm range nhập cho tháng hiện hành
				now.minusMonths(2);
				final var ymS = now.minusMonths(2);
				final var ymE = now;

				standardizedRangeInput(sheet, formulaEvaluator, ymS, ymE);

				// Thêm dòng trống để điền thông tin mới
				addRowForInsertData(ankenNoUpdate, sheet, l, formulaEvaluator);

				addRowForInsertData(ankenNoInsert, sheet, l, formulaEvaluator);

				System.out.println("Add row created successfully.");

				// Cập nhật lại công thức
				updatedTotalActualHoursFormula(sheet, formulaEvaluator);
				updatedTotalFooterFormula(sheet, formulaEvaluator);

				// Chạy lại toàn bộ công thức
				evaluateAllFormula(workbook);

				// create
				// -- create new sheet ? tự tạo template
				// update
			}
			// Ghi dữ liệu vào tệp tin
			saveToNewFileSchedule(workbook, "03006523", null);

		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

	private void fillRowForInput(final List<BacklogDetail> backlogs, final Sheet sheet,
			final FormulaEvaluator formulaEvaluator) {
		final Map<String, List<BacklogDetail>> groupedBacklogs = backlogs.stream()
				.collect(Collectors.groupingBy(BacklogDetail::getAnkenNo));
		for (final Map.Entry<String, List<BacklogDetail>> entry : groupedBacklogs.entrySet()) {
			final var ankenNo = entry.getKey();
			var curBacklog = entry.getValue();
			addRowForInsertData(ankenNo, sheet, curBacklog, formulaEvaluator);
		}
	}

	private void fillBacklogInfo(String projecCd, final List<PjjyujiDetail> pds, final List<BacklogDetail> bds,
			Workbook workbook) {
		final var yearMonths = pds.stream().map(PjjyujiDetail::getTargetYmd).map(YearMonth::from).distinct().toList();
		final var now = YearMonth.now();
		final var ymS = yearMonths.stream().min(YearMonth::compareTo).orElse(now);
		final var ymE = yearMonths.stream().max(YearMonth::compareTo).orElse(now);

		final var backlogBug = bds.stream() //
				.filter(x -> StringUtils.equals(issuTypeBug, x.getIssueType())) //
				.toList();
		final var backlogSpec = bds.stream() //
				.filter(x -> StringUtils.equals(issuTypeSpec, x.getIssueType())) //
				.toList();
		final var backlogPg = bds.stream() //
				.filter(x -> !StringUtils.equals(issuTypeSpec, x.getIssueType())
						&& !StringUtils.equals(issuTypeBug, x.getIssueType())) //
				.toList();
		final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		final var sheetIterator = workbook.sheetIterator();
		while (sheetIterator.hasNext()) {
			final var sheet = sheetIterator.next();
			// Kiểm tra là sheet điền schedule
			if (!ScheduleHelper.isScheduleSheet(sheet)) {
				continue;
			}
			final var sheetName = StringUtils.lowerCase(sheet.getSheetName());
			System.out.println("Sheet: " + sheetName);

			standardizedRangeInput(sheet, formulaEvaluator, ymS, ymE);

			if (StringUtils.equals(sheetName, "pg_spec")) {
				fillRowForInput(backlogSpec, sheet, formulaEvaluator);
			} else if (StringUtils.equals(sheetName, "pg_bug")) {
				fillRowForInput(backlogBug, sheet, formulaEvaluator);
			} else {
				fillRowForInput(backlogPg, sheet, formulaEvaluator);
			}

			// Cập nhật lại công thức
			updatedTotalActualHoursFormula(sheet, formulaEvaluator);

			updatedTotalFooterFormula(sheet, formulaEvaluator);

			// Chạy lại toàn bộ công thức
			evaluateAllFormula(workbook);
		}
	}

	/**
	 *
	 * @param projecCd
	 * @param datas
	 * @param workbook
	 */
	private void fillScheduleInfo(String projecCd, final List<PjjyujiDetail> pds, final List<BacklogDetail> bds,
			Workbook workbook) {

		fillBacklogInfo(projecCd, pds, bds, workbook);

		// Chạy lại toàn bộ công thức
		evaluateAllFormula(workbook);
	}

	public void createSchedule(String projecCd, Path backlogSchedulePath, final List<PjjyujiDetail> pds,
			final List<BacklogDetail> bds) {
		try (var fis = BacklogExcelUtil.class.getClassLoader().getResourceAsStream(scheduleTemplatePath);
				Workbook workbook = new XSSFWorkbook(fis)) {

			fillScheduleInfo(projecCd, pds, bds, workbook);

			// Ghi dữ liệu vào tệp tin mới
			saveToNewFileSchedule(workbook, projecCd, backlogSchedulePath);

		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Sau khi xử lý xong schedule thì thực hiện ghi vào file mới
	 *
	 * @param workbook
	 * @param pjCd
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private void saveToNewFileSchedule(Workbook workbook, String pjCd, Path scheduleFolerPath) throws IOException {
		// Ghi dữ liệu vào tệp tin
		final var fileName = StringUtils.replaceEach(templateFile, new String[] { "{projectCd}", },
				new String[] { pjCd });
		// new file schedule
		File targetFile = null;
		if (scheduleFolerPath != null) {
			final var filePath = scheduleFolerPath.resolve(fileName);
			// Convert the Path object to a File object
			targetFile = filePath.toFile();
		} else {
			targetFile = new File(fileName);
		}

		try (var fileOut = new FileOutputStream(targetFile, false);) {

			workbook.write(fileOut);
			System.out.println("New file created: " + targetFile);
		}
	}

	private void evaluateAllFormula(Workbook workbook) {
		// Create a formula evaluator
		final var evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		// Update all formulas in the sheet
		evaluator.evaluateAll();

	}
}

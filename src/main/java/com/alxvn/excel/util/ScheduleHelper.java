/**
 *
 */
package com.alxvn.excel.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author KEDD
 *
 */
public class ScheduleHelper {

	private static final Logger log = LoggerFactory.getLogger(ScheduleHelper.class);

	public static final String TOTAL = "TOTAL";
	public static final String SCHE_ROW_TITLE_IDX = "SCHE_ROW_TITLE_IDX";
	public static final DecimalFormat df = new DecimalFormat("#.##");

	public static final String SCHE_PROJECT_CODE = "SCHE_PROJECT_CODE";
	public static final String SCHE_PROJECT_CODE_JP = "SCHE_PROJECT_CODE_JP";
	/* Sheet */
	public static final String SCHE_SHEET_PIC_IDX = "SCHE_SHEET_PIC_IDX";
	public static final String SCHE_SHEET_PROCESS_IDX = "SCHE_SHEET_PROCESS_IDX";
	public static final String SCHE_SHEET_ROW_TITLE_IDX = "SCHE_SHEET_ROW_TITLE_IDX";
	public static final String SCHE_SHEET_TICKET_IDX = "SCHE_SHEET_TICKET_IDX";
	public static final String SCHE_SHEET_TOTAL_IDX = "SCHE_SHEET_TOTAL_IDX";

	public static final String SCHE_SHEET_EXPECTED_HOURS_IDX = "SCHE_SHEET_EXPECTED_HOURS_IDX";
	public static final String SCHE_SHEET_EXPECTED_BEGIN_IDX = "SCHE_SHEET_EXPECTED_BEGIN_IDX";
	public static final String SCHE_SHEET_EXPECTED_END_IDX = "SCHE_SHEET_EXPECTED_END_IDX";
	public static final String SCHE_SHEET_EXPECTED_DELIVERY_IDX = "SCHE_SHEET_EXPECTED_DELIVERY_IDX";

	public static final String SCHE_SHEET_STATUS_IDX = "SCHE_SHEET_STATUS_IDX";
	public static final String SCHE_SHEET_PROGRESS_IDX = "SCHE_SHEET_PROGRESS_IDX";

	public static final String SCHE_SHEET_ACTUAL_HOURS_IDX = "SCHE_SHEET_ACTUAL_HOURS_IDX";
	public static final String SCHE_SHEET_ACTUAL_BEGIN_IDX = "SCHE_SHEET_ACTUAL_BEGIN_IDX";
	public static final String SCHE_SHEET_ACTUAL_END_IDX = "SCHE_SHEET_ACTUAL_END_IDX";
	public static final String SCHE_SHEET_ACTUAL_DELIVERY_IDX = "SCHE_SHEET_ACTUAL_DELIVERY_IDX";

	private static final DateTimeFormatter FORMATTER_YYYYMMDD = DateTimeFormatter.ofPattern("yyyy/MM/dd");

	private ScheduleHelper() {
		throw new IllegalStateException("Utility class");
	}

	public static boolean isTotalRow(Row row, String totalVal, Integer scheTotalIdx) {
		final Cell cell = CellUtil.getCell(row, scheTotalIdx);
		final String val = getCellValueAsString(cell);
		return StringUtils.equals(val, totalVal);

	}

	public static boolean isValidSchFileName(File fileSchedule) {
		if (Objects.nonNull(fileSchedule)) {
			final String fileName = fileSchedule.getName();
			final String pattern = "^(QDA-0222a_プロジェクト管理表_).*\\.(xls|xlsx|xlsm)$";
			final Pattern r = Pattern.compile(pattern, Pattern.CANON_EQ);
			final Matcher matcher = r.matcher(fileName);
			if (matcher.find()) {
				log.debug("isValidScheduleFile.isValid_File_Name: {}", fileSchedule);
				return true;
			}
		}
		log.debug("isValidScheduleFile.isNotValid_File_Name: {}", fileSchedule);
		return false;
	}

	public static boolean isValidScheduleFileName(String fileName) {
		if (Objects.nonNull(fileName)) {
			final String pattern = "^(QDA-0222a_プロジェクト管理表_).*\\.(xls|xlsx|xlsm)$";
			final Pattern r = Pattern.compile(pattern, Pattern.CANON_EQ);
			final Matcher matcher = r.matcher(fileName);
			if (matcher.find()) {
				log.debug("isValidScheduleFile.isValid_File_Name: {}", fileName);
				return true;
			}
		}
		log.debug("isValidScheduleFile.isNotValid_File_Name: {}", fileName);
		return false;
	}

	public static boolean isScheduleHaveValue(File fileSchedule, YearMonth targetYm, Predicate<Sheet> isScheduleSheet,
			Predicate<Row> isTotalRow) {
		if (Objects.isNull(fileSchedule)) {
			log.debug("isScheduleHaveValue - fileSchedule is null !!!");
			return false;
		}
		final String fileName = fileSchedule.getName();
		if (Objects.isNull(targetYm)) {
			log.debug("isScheduleHaveValue - targetYm is null: {}", fileName);
			return false;
		}
		if (!isValidSchFileName(fileSchedule)) {
			log.debug("isScheduleHaveValue.isNotValid: {}", fileName);
			return false;
		}

		try (Workbook workbook = WorkbookFactory.create(fileSchedule)) {
			if (Objects.isNull(workbook)) {
				log.debug("isScheduleHaveValue.isNotValid: {}", fileName);
				return false;
			}
			final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			for (final Sheet sheet : workbook) {
				if (!isScheduleSheet.test(sheet)) {
					continue;
				}
				final Row schTitleRow = getScheTitleRow(sheet);
				// have date of target month(does not need fully date)
				final Pair<Integer, Integer> colStartEnd = findRangeColIndex(evaluator, sheet, targetYm);
				if (Objects.isNull(colStartEnd)) {
					continue;
				}
				final Integer scheColStartDate = colStartEnd.getLeft();
				final Integer scheColEndDate = colStartEnd.getRight();
				// Check range input have value
				boolean isInputed = false;
				final Iterator<Row> rowIterator = sheet.rowIterator();
				while (rowIterator.hasNext()) {
					final Row row = rowIterator.next();
					if (row == null || row.getRowNum() <= schTitleRow.getRowNum()) {
						continue;
					}
					if (isTotalRow.test(row)) {
						break;
					}
					final Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						final Cell c = cellIterator.next();
						if (c == null) {
							// The spreadsheet is empty in this cell
							continue;
						}
						final int cn = c.getColumnIndex();
						// ignore if not in range
						if (cn < scheColStartDate) {
							continue;
						}
						if (cn > scheColEndDate) {
							break;
						}
						// check if range have inputed
						if (c.getCellType() != CellType.BLANK) {
							isInputed = true;
							break;
						}

					}
					if (isInputed) {
						break;
					}
				}
				if (!isInputed) {
					log.debug("Workbook {} at sheet {}  không có thông tin !", fileName, sheet.getSheetName());
					continue;
				}
				log.debug("isValidScheduleFile.isValid: {} - {}", fileSchedule.getName(), sheet.getSheetName());
				return true;
			}
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}
		log.debug("isScheduleHaveValue.isNotValid: {}", fileSchedule.getName());
		return false;
	}

	public static boolean isScheduleHaveValue1(File fileSchedule, YearMonth targetYm, Predicate<Sheet> isScheduleSheet,
			Predicate<Row> isTotalRow) {
		if (Objects.isNull(fileSchedule)) {
			log.debug("isValidScheduleFile - fileSchedule is null !!!");
			return false;
		}
		final String fileName = fileSchedule.getName();
		if (Objects.isNull(targetYm)) {
			log.debug("isValidScheduleFile - targetYm is null: {}", fileName);
			return false;
		}
		if (!isValidSchFileName(fileSchedule)) {
			log.debug("isValidScheduleFile.isNotValid: {}", fileName);
			return false;
		}

		try (Workbook workbook = WorkbookFactory.create(fileSchedule)) {
			if (Objects.isNull(workbook)) {
				log.debug("isValidScheduleFile.isNotValid: {}", fileName);
				return false;
			}
			final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				final Sheet sheet = workbook.getSheetAt(i);
				if (!isScheduleSheet.test(sheet)) {
					continue;
				}
				final Row schTitleRow = getScheTitleRow(sheet);
				// have date of target month(does not need fully date)
				final Pair<Integer, Integer> colStartEnd = findRangeColIndex(evaluator, sheet, targetYm);
				if (Objects.isNull(colStartEnd)) {
					continue;
				}
				final Integer scheColStartDate = colStartEnd.getLeft();
				final Integer scheColEndDate = colStartEnd.getRight();
				// Check range input have value
				boolean isInputed = false;
				for (final Row row : sheet) {
					if (row == null || row.getRowNum() <= schTitleRow.getRowNum()) {
						continue;
					}
					if (isTotalRow.test(row)) {
						break;
					}
					for (int cn = scheColStartDate; cn <= scheColEndDate; cn++) {
						final Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
						if (c != null && c.getCellType() != CellType.BLANK) {
							isInputed = true;
							break;
						}
					}
				}
				if (!isInputed) {
					log.debug("Workbook {} at sheet {}  không có thông tin !", fileName, sheet.getSheetName());
					continue;
				}
				log.debug("isValidScheduleFile.isValid: {} - {}", fileSchedule.getName(), sheet.getSheetName());
				return true;
			}
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}
		log.debug("isValidScheduleFile.isNotValid: {}", fileSchedule.getName());
		return false;
	}

	public boolean isMergedCell(Sheet sheet, int row, int column) {
		for (final CellRangeAddress range : sheet.getMergedRegions()) {
			if (range.isInRange(row, column)) {
				return true;
			}
		}
		return false;
	}

	public static CellRangeAddress getMergedRegionForCell(Cell c) {
		final Sheet s = c.getRow().getSheet();
		for (final CellRangeAddress mergedRegion : s.getMergedRegions()) {
			if (mergedRegion.isInRange(c.getRowIndex(), c.getColumnIndex())) {
				// This region contains the cell in question
				return mergedRegion;
			}
		}
		// Not in any
		return null;
	}

	public static Cell findCellScheduleDelivery(Sheet sheet) {
		// 予定
		for (final Row row : sheet) {
			if (row == null) {
				continue;
			}
			final int lastColumn = Math.max(row.getLastCellNum(), 1);
			for (int cn = 0; cn < lastColumn; cn++) {
				final Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				if (c == null) {
					// The spreadsheet is empty in this cell
				} else {
					// Do something useful with the cell's contents
					final String val = getCellValueAsString(c);
					if (StringUtils.contains(val, "納品日")) {
						return c;
					}
				}
			}
		}
		return null;
	}

	public static String readContentCell(Sheet sheet, Cell cell) {
		final CellRangeAddress cra = getMergedRegionForCell(cell);
		if (cra != null) {
			return readContentFromMergedCells(sheet, cra);
		}
		return getCellValueAsString(cell);
	}

	public static String readContentFromMergedCells(Sheet sheet, CellRangeAddress mergedCells) {

		if (Objects.isNull(mergedCells)) {
			return null;
		}
		final Cell cell = sheet.getRow(mergedCells.getFirstRow()).getCell(mergedCells.getFirstColumn());
		return getCellValueAsString(cell);
	}

	private static String cellColName(Integer columnIndex) {
		return CellReference.convertNumToColString(columnIndex);
	}

	@Deprecated
	public static Pair<Integer, Integer> findScheColByTarget(Sheet sheet, YearMonth targetMonth) {
		log.debug("findScheColByTarget: {} - {}", sheet.getSheetName(), targetMonth);
		if (Objects.isNull(targetMonth)) {
			return null;
		}
		final FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		final LocalDate start = targetMonth.atDay(1);
		final LocalDate end = targetMonth.atEndOfMonth();
		final int diffInMonth = (int) Math.abs(ChronoUnit.DAYS.between(start, end));
		final Row scheRowTitle = getScheTitleRow(sheet);
		final int lastColumn = Math.max(scheRowTitle.getLastCellNum(), 1);
		Integer colStartIdx = null;
		for (int cn = 0; cn < lastColumn; cn++) {
			final Cell c = scheRowTitle.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
			if (c == null) {
				continue;
			}
			try {
				// skip n days
				final CellValue cellValue = formulaEvaluator.evaluate(c);
				final LocalDate date = getTargetYmd(c, cellValue);
				if (Objects.equals(date, start)) {
					colStartIdx = cn;
					break;
				}
			} catch (final Exception e) {
			}
		}
		if (Objects.isNull(colStartIdx)) {
			log.debug("findScheColByTarget.warn.can't detect range: {}", sheet.getSheetName());
			return null;
		}
		Integer colStart = null;
		Integer colEnd = null;
		final Cell cellStart = scheRowTitle.getCell(colStartIdx);
		if (Objects.isNull(cellStart)) {
			log.debug("findScheColByTarget.warn.can't detect range: {}", sheet.getSheetName());
			return null;
		}
		final LocalDate dateStart = getTargetYmd(cellStart, formulaEvaluator.evaluate(cellStart));
		if (Objects.equals(dateStart, start)) {
			colStart = cellStart.getColumnIndex();
		}
		final Cell cellEnd = scheRowTitle.getCell(colStartIdx + diffInMonth);
		if (Objects.isNull(cellEnd)) {
			log.debug("findScheColByTarget.warn.can't detect range: {}", sheet.getSheetName());
			return null;
		}
		final LocalDate dateEnd = getTargetYmd(cellEnd, formulaEvaluator.evaluate(cellEnd));
		if (Objects.equals(dateEnd, end)) {
			colEnd = cellEnd.getColumnIndex();
		}

		if (Objects.isNull(colStart) || Objects.isNull(colEnd)) {
			log.debug("findScheColByTarget.warn.can't detect range: {}", sheet.getSheetName());
			return null;
		}
		final String sCol = cellColName(colStart);
		final String eCol = cellColName(colEnd);
		log.debug("findScheColByTarget.isvalid: {} -- {} - {}", sheet.getSheetName(), sCol, eCol);
		return Pair.of(colStart, colEnd);

	}

	public static Pair<Integer, Integer> findRangeColIndex(FormulaEvaluator evaluator, Sheet sheet,
			YearMonth targetMonth) {
		log.debug("findRangeColIndex: {} - {}", sheet.getSheetName(), targetMonth);
		if (Objects.isNull(targetMonth)) {
			return null;
		}

		final Row schTitleRow = getScheTitleRow(sheet);

		// Check have row of target month
		Integer colStart = null;
		Integer colEnd = null;

		final Iterator<Cell> cellIterator = schTitleRow.cellIterator();
		while (cellIterator.hasNext()) {
			final Cell c = cellIterator.next();
			if (c == null) {
				// The spreadsheet is empty in this cell
				continue;
			}

			// TODO: using DataFormatter
//			final DataFormatter dataFormatter = new DataFormatter();
//			final String cellValue = dataFormatter.formatCellValue(c, evaluator);
//			final LocalDate date = LocalDate.parse(cellValue, FORMATTER_YYYYMMDD);
//			final YearMonth yearMonth = YearMonth.from(date);

			try {
				final YearMonth yearMonth = YearMonth
						.from(LocalDate.parse(getCellValueAsString(c, evaluator.evaluate(c)), FORMATTER_YYYYMMDD));
				final int comparison1 = yearMonth.compareTo(targetMonth);
				if (comparison1 == 0) {
					if (colStart == null) {
						colStart = c.getColumnIndex(); // first cell match target month
					}
					if (colEnd == null || colEnd < c.getColumnIndex()) {
						colEnd = c.getColumnIndex();
					}
				}
				if (comparison1 > 0) {
					break;
				}
			} catch (final Exception e) {
				// Handle exception or log error if necessary
			}
		}

		if (Objects.isNull(colStart) || Objects.isNull(colEnd)) {
			log.debug("findRangeColIndex.warn.can't detect range: {}", sheet.getSheetName());
			return null;
		}

		final String sCol = cellColName(colStart);
		final String eCol = cellColName(colEnd);
		log.debug("findRangeColIndex.isvalid: {} -- {} - {}", sheet.getSheetName(), sCol, eCol);
		return Pair.of(colStart, colEnd);
	}

	public static Pair<Integer, Integer> findRangeColIndex1(FormulaEvaluator evaluator, Sheet sheet,
			YearMonth targetMonth) {
		log.debug("findRangeColIndex: {} - {}", sheet.getSheetName(), targetMonth);
		if (Objects.isNull(targetMonth)) {
			return null;
		}
		final Row schTitleRow = getScheTitleRow(sheet);
		final int lastColumn = Math.max(schTitleRow.getLastCellNum(), 1);
		// Check have row of target month
		Integer colStart = null;
		Integer colEnd = null;
		for (int cn = 0; cn < lastColumn; cn++) {
			final Cell c = schTitleRow.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
			if (c == null) {
				// The spreadsheet is empty in this cell
				continue;
			}
			// Do something useful with the cell's contents
			try {
				final YearMonth yearMonth = YearMonth
						.from(LocalDate.parse(getCellValueAsString(c, evaluator.evaluate(c)), FORMATTER_YYYYMMDD));
				final int comparison1 = yearMonth.compareTo(targetMonth);
				if (comparison1 == 0) {
					if (colStart == null) {
						colStart = c.getColumnIndex(); // first cell match target month
					}
					if (colEnd == null || colEnd < c.getColumnIndex()) {
						colEnd = c.getColumnIndex();
					}
				}
				if (comparison1 > 0) {
					break;
				}
			} catch (final Exception e) {
			}
		}

		if (Objects.isNull(colStart) || Objects.isNull(colEnd)) {
			log.debug("findRangeColIndex.warn.can't detect range: {}", sheet.getSheetName());
			return null;
		}
		final String sCol = cellColName(colStart);
		final String eCol = cellColName(colEnd);
		log.debug("findRangeColIndex.isvalid: {} -- {} - {}", sheet.getSheetName(), sCol, eCol);
		return Pair.of(colStart, colEnd);

	}

	public static String getPjCd(Sheet sheet) {
		for (final Row row : sheet) {
			if (row == null) {
				continue;
			}
			final int lastColumn = Math.max(row.getLastCellNum(), 1);
			for (int cn = 0; cn < lastColumn; cn++) {
				final Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				if (c == null) {
					// The spreadsheet is empty in this cell
					continue;
				}
				// Do something useful with the cell's contents
				final String val = getCellValueAsString(c);
				if (StringUtils.contains(val, "PJCODE")) {
					final Cell nextCell = row.getCell(cn + 1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					return StringUtils.trim(getCellValueAsString(nextCell));
				}
			}
		}
		return StringUtils.EMPTY;
	}

	public static String getPjNo1(Sheet sheet) {
		for (final Row row : sheet) {
			if (row == null) {
				continue;
			}

			final int lastColumn = Math.max(row.getLastCellNum(), 1);
			for (int cn = 0; cn < lastColumn; cn++) {
				final Cell cell = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				if (cell == null) {
					// The spreadsheet is empty in this cell
					continue;
				}

				final DataFormatter dataFormatter = new DataFormatter();
				final String cellValue = dataFormatter.formatCellValue(cell);

				if (StringUtils.containsIgnoreCase(cellValue, "PJ-NO")) {
					final Cell nextCell = row.getCell(cn + 1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					return StringUtils.trim(dataFormatter.formatCellValue(nextCell));
				}
			}
		}

		throw new IllegalArgumentException("Can't detect Project No !!!");
	}

	public static String getPjNo(Sheet sheet) {
		for (final Row row : sheet) {
			if (row == null) {
				continue;
			}
			final int lastColumn = Math.max(row.getLastCellNum(), 1);
			for (int cn = 0; cn < lastColumn; cn++) {
				final Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				if (c == null) {
					// The spreadsheet is empty in this cell
					continue;
				}
				// Do something useful with the cell's contents
				final String val = getCellValueAsString(c);
				if (StringUtils.contains(val, "PJ-NO")) {
					final Cell nextCell = row.getCell(cn + 1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					return StringUtils.trim(getCellValueAsString(nextCell));
				}
			}
		}
		throw new IllegalArgumentException("Can't detect Project No !!!");
	}

	public static String getPjNoEc(Sheet sheet) {
		for (final Row row : sheet) {
			if (row == null) {
				continue;
			}
			final int lastColumn = Math.max(row.getLastCellNum(), 1);
			for (int cn = 0; cn < lastColumn; cn++) {
				final Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
				if (c == null) {
					// The spreadsheet is empty in this cell
					continue;
				}
				// Do something useful with the cell's contents
				final String val = getCellValueAsString(c);
				if (StringUtils.contains(val, "PJ-NO")) {
					final Pattern pattern = Pattern.compile("(\\d+)");
					final Matcher matcher = pattern.matcher(val);
					if (matcher.find()) {
						return StringUtils.defaultString(matcher.group(1));
					}
				}
			}
		}
		throw new IllegalArgumentException("Can't detect Project No !!!");
	}

	public static Row getScheTitleRow(Sheet sheet) {
		for (final Row row : sheet) {
			if (row == null) {
				continue;
			}
			final Cell cell = CellUtil.getCell(row, CellReference.convertColStringToIndex("A"));
			final String no = getCellValueAsString(cell);
			if (StringUtils.equals(no, "No")) {
				return row;
			}
		}
		throw new IllegalArgumentException("Sheet " + sheet.getSheetName() + " khong co dong title");
	}

	public static String resolveProjectCode(String content, String sheetName) {
		final Pattern pattern = Pattern.compile("(\\d+)");
		final Matcher matcher = pattern.matcher(content);
		if (matcher.find()) {
			return StringUtils.trim(StringUtils.defaultString(matcher.group(1)));
		}
		log.debug("resolveProjectCode.Khong lay duoc project code: {}", sheetName);
		return StringUtils.EMPTY;
	}

	public static boolean isScheduleSheet(Sheet sheet) throws EncryptedDocumentException {
		if (Objects.isNull(sheet)) {
			return false;
		}
		final Pattern r = Pattern.compile("^(開発|開発スケジュール|PG|Acceptance|PD|P).*$", Pattern.CANON_EQ);
		final Matcher matcher = r.matcher(sheet.getSheetName());
		return matcher.find();
	}

	private static String resolveDate(Cell cell) {
		// Get the date value from the cell
		final Date date = cell.getDateCellValue();
		// Convert the Date to Instant
		final Instant instant = date.toInstant();

		// Convert the Instant to LocalDateTime
		final LocalDateTime localDateTime = LocalDateTime.ofInstant(instant, ZoneId.systemDefault());

		// Format the LocalDateTime to a string
		return localDateTime.format(FORMATTER_YYYYMMDD);
	}

	/**
	 * This method for the type of data in the cell, extracts the data and returns
	 * it as a string.
	 */
	private static String getCellValueForFormula(final CellType cellType, final Cell cell) {
		String strCellValue = null;
		if (Objects.nonNull(cell)) {
			if (cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
				return resolveDate(cell);
			}
			switch (cell.getCellType()) {
			case STRING:
				strCellValue = cell.toString();
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					strCellValue = resolveDate(cell);
				} else {
					final Double value = cell.getNumericCellValue();
					strCellValue = String.valueOf(value);
				}
				break;
			case BOOLEAN:
				strCellValue = String.valueOf(cell.getBooleanCellValue());
				break;
			case BLANK:
				strCellValue = "";
				break;
			default:
				break;
			}
		}
		return strCellValue;
	}

	public static String getCellValueAsString(Cell cell, CellValue cellValue) {
		if (Objects.nonNull(cellValue)) {
			switch (cellValue.getCellType()) {
			case BOOLEAN:
				return String.valueOf(cellValue.getBooleanValue());
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return resolveDate(cell);
				} else {
					final Double value = cell.getNumericCellValue();
					return String.valueOf(value);
				}
			case STRING:
				return cellValue.getStringValue();
			case FORMULA:
				return getCellValueForFormula(cellValue.getCellType(), cell);
			case BLANK:
				return "";
			case ERROR:
				return FormulaError.forInt(cellValue.getErrorValue()).getString();
			default:
				return "";
			}
		}
		return getCellValueAsString(cell);
	}

	public static String getCellValueAsString(Cell cell) {
		String strCellValue = null;
		if (Objects.nonNull(cell)) {
			switch (cell.getCellType()) {
			case STRING:
				strCellValue = cell.toString();
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					strCellValue = resolveDate(cell);
				} else {
					final Double value = cell.getNumericCellValue();
					strCellValue = String.valueOf(value);
				}
				break;
			case BOOLEAN:
				strCellValue = String.valueOf(cell.getBooleanCellValue());
				break;
			case BLANK:
				strCellValue = "";
				break;
			default:
				break;
			}
		}
		return strCellValue;
	}


	public static String createFolderBackup(String targetExecutePath) {
		final String folderPath = targetExecutePath + File.separator + "ScheBackup";
		return createDirectories(folderPath);
	}

	public static String createDirectories(String folderPath) {

		final Path dirPath = Paths.get(folderPath);
		try {
			Files.createDirectories(dirPath);
		} catch (final IOException e) {
			throw new IllegalArgumentException("Error creating download folder:", e);
		}

		if (Files.exists(dirPath)) {
			try {
				FileUtils.cleanDirectory(dirPath.toFile());
			} catch (final IOException e) {
				throw new IllegalArgumentException("Error cleaning download folder:", e);
			}
		}

		return folderPath;
	}

	public static String createFolderDownload(String targetExecutePath) {
		final String folderPath = targetExecutePath + File.separator + "ScheDownload";
		return createDirectories(folderPath);
	}

	public static String createFolderResult(String targetExecutePath) {
		final String folderPath = targetExecutePath + File.separator + "Result";
		return createDirectories(folderPath);
	}


	public static LocalDate getTargetYmd(Cell cell, CellValue cellValue) {
		final String date = getCellValueAsString(cell, cellValue);
		return LocalDate.parse(date, FORMATTER_YYYYMMDD);
	}

	public static LocalDate getTargetYmdEc(Sheet sheet, Cell cell, Row scheRowTitle) {
		final FormulaEvaluator evaluator = scheRowTitle.getSheet().getWorkbook().getCreationHelper()
				.createFormulaEvaluator();
		final CellRangeAddress cra = getMergedRegionForCell(
				sheet.getRow(scheRowTitle.getRowNum() - 1).getCell(cell.getColumnIndex()));
		if (Objects.nonNull(cra)) {
			final CellValue cellValue = evaluator.evaluate(cell);
			final String date = getCellValueAsString(cell, cellValue);
			final String targetYmd = readContentFromMergedCells(sheet, cra);
			final LocalDate targetDate = LocalDate.parse(targetYmd, FORMATTER_YYYYMMDD);
			return targetDate.plusDays(Integer.parseInt(date)).minusDays(1);
		}

		throw new IllegalArgumentException("Can't detect target date !!!");
	}

	public static String convertColumnIndexToName(int columnIndex) {
		final StringBuilder columnName = new StringBuilder();

		while (columnIndex >= 0) {
			final int remainder = columnIndex % 26;
			columnName.insert(0, (char) ('A' + remainder));
			columnIndex = columnIndex / 26 - 1;
		}

		return columnName.toString();
	}
}

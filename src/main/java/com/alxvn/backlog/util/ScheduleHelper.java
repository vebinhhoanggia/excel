/**
 *
 */
package com.alxvn.backlog.util;

import java.io.File;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alxvn.backlog.dto.PjjyujiDetail;

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

	public static boolean isTotalRow(final Row row, final String totalVal, final Integer scheTotalIdx) {
		final var cell = CellUtil.getCell(row, scheTotalIdx);
		final var val = getCellValueAsString(cell);
		return StringUtils.equals(val, totalVal);

	}

	public static boolean isValidSchFileName(final File fileSchedule) {
		if (Objects.nonNull(fileSchedule)) {
			final var fileName = fileSchedule.getName();
			final var pattern = "^(QDA-0222a_プロジェクト管理表_).*\\.(xls|xlsx|xlsm)$";
			final var r = Pattern.compile(pattern, Pattern.CANON_EQ);
			final var matcher = r.matcher(fileName);
			if (matcher.find()) {
				log.debug("isValidScheduleFile.isValid_File_Name: {}", fileSchedule);
				return true;
			}
		}
		log.debug("isValidScheduleFile.isNotValid_File_Name: {}", fileSchedule);
		return false;
	}

	public static boolean isValidScheduleFileName(final String fileName) {
		if (Objects.nonNull(fileName)) {
			final var pattern = "^(QDA-0222a_プロジェクト管理表_).*\\.(xls|xlsx|xlsm)$";
			final var r = Pattern.compile(pattern, Pattern.CANON_EQ);
			final var matcher = r.matcher(fileName);
			if (matcher.find()) {
				log.debug("isValidScheduleFile.isValid_File_Name: {}", fileName);
				return true;
			}
		}
		log.debug("isValidScheduleFile.isNotValid_File_Name: {}", fileName);
		return false;
	}

	public static boolean isScheduleSheet(final Sheet sheet) throws EncryptedDocumentException {
		if (Objects.isNull(sheet)) {
			return false;
		}
		final var r = Pattern.compile("^(開発|開発スケジュール|PG|Acceptance|PD|P).*$", Pattern.CANON_EQ);
		final var matcher = r.matcher(sheet.getSheetName());
		return matcher.find();
	}

	public static String resolveDate(final Cell cell) {
		// Get the date value from the cell
		final var date = cell.getDateCellValue();
		// Convert the Date to Instant
		final var instant = date.toInstant();

		// Convert the Instant to LocalDateTime
		final var localDateTime = LocalDateTime.ofInstant(instant, ZoneId.systemDefault());

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

	public static String getCellValueAsString(final Cell cell, final CellValue cellValue) {
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

	public static String getCellValueAsString(final Cell cell) {
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

	public static Double getWrMinute(final List<PjjyujiDetail> wrDetails, final String ankenNo) {
		return Double.parseDouble(df.format(wrDetails.stream().filter(x -> StringUtils.equals(ankenNo, x.getAnkenNo()))
				/**/
				.map(PjjyujiDetail::getMinute).map(Double::valueOf).reduce(0.0, Double::sum)));
	}

	public static String convertColumnIndexToName(int columnIndex) {
		final var columnName = new StringBuilder();

		while (columnIndex >= 0) {
			final var remainder = columnIndex % 26;
			columnName.insert(0, (char) ('A' + remainder));
			columnIndex = columnIndex / 26 - 1;
		}

		return columnName.toString();
	}

	public static CellRangeAddress getMergedRegionForCell(final Cell c) {
		final var s = c.getRow().getSheet();
		for (final CellRangeAddress mergedRegion : s.getMergedRegions()) {
			if (mergedRegion.isInRange(c.getRowIndex(), c.getColumnIndex())) {
				// This region contains the cell in question
				return mergedRegion;
			}
		}
		// Not in any
		return null;
	}

	public static String readContentFromMergedCells(final Sheet sheet, final CellRangeAddress mergedCells) {

		if (Objects.isNull(mergedCells)) {
			return null;
		}
		final var cell = sheet.getRow(mergedCells.getFirstRow()).getCell(mergedCells.getFirstColumn());
		return getCellValueAsString(cell);
	}

	public static String readContentCell(final Sheet sheet, final Cell cell) {
		final var cra = getMergedRegionForCell(cell);
		if (cra != null) {
			return readContentFromMergedCells(sheet, cra);
		}
		return getCellValueAsString(cell);
	}

	public static LocalDate getTargetYmd(final Cell cell, final CellValue cellValue) {
		final var date = getCellValueAsString(cell, cellValue);
		return LocalDate.parse(date, FORMATTER_YYYYMMDD);
	}
}

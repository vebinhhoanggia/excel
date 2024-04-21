/**
 *
 */
package com.alxvn.backlog.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;

import com.alxvn.backlog.handle.IncorrectFullNameException;

/**
 * @author KEDD
 *
 */
public class WorkingReportHelper {

	private WorkingReportHelper() {
		throw new IllegalStateException("Utility class");
	}

	public static String getMailIdFromFullName(String fullName) throws IncorrectFullNameException {
		final var blank = " ";
		if (StringUtils.isNotBlank(fullName)) {
			final var builder = new StringBuilder();

			final var lastIndexOf = fullName.lastIndexOf(blank);
			if (lastIndexOf < 0) {
				throw new IncorrectFullNameException("Can not determin mail id: " + fullName);
			}
			final var name = fullName.substring(lastIndexOf + 1, fullName.length());
			builder.append(name);

			final var surname = fullName.substring(0, lastIndexOf);
			final var splited = surname.split(blank);
			for (final String str : splited) {
				final var string = String.valueOf(Character.toLowerCase(str.charAt(0)));
				builder.append(string);
			}
			return builder.toString();
		}
		return "";
	}

	public static Integer sumMinute(final String minute1, final String minute2, final String minute3,
			final String minute4, final String minute5) {
		return NumberUtils.toInt(minute1, 0) + NumberUtils.toInt(minute2, 0) + NumberUtils.toInt(minute3, 0)
				+ NumberUtils.toInt(minute4, 0) + NumberUtils.toInt(minute5, 0);
	}
}

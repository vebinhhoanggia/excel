/**
 *
 */
package com.alxvn.backlog.util;

import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;

/**
 * @author KEDD
 *
 */
public class Helper {

	public static String getAnkenNo(String content) {
		// use for redmine #xxxx
		final var redminePattern = "#\\d+";
		final var patternRedmine = Pattern.compile(redminePattern);
		final var matcherRedmine = patternRedmine.matcher(content);
		if (matcherRedmine.find()) {
			return StringUtils.defaultString(matcherRedmine.group(0));
		}
		// use for backlog sym
		final var backlogSymPattern = "SYMPHO-\\d+";
		final var patternBacklog = Pattern.compile(backlogSymPattern);
		final var matcherSymBacklog = patternBacklog.matcher(content);
		if (matcherSymBacklog.find()) {
			return StringUtils.defaultString(matcherSymBacklog.group(0));
		}
		// use for backlog ifront
		final var backlogifrontPattern = "IFRONT-\\d+";
		final var patternIFrontBacklog = Pattern.compile(backlogifrontPattern);
		final var matcherIfrontBacklog = patternIFrontBacklog.matcher(content);
		if (matcherIfrontBacklog.find()) {
			return StringUtils.defaultString(matcherIfrontBacklog.group(0));
		}

		// using redmine
		final var specStr = "#";
		final var pattern = Pattern.compile("(\\#\\d+|\\d+)");
		final var matcher = pattern.matcher(content);
		if (matcher.find()) {
			final var contentStr = StringUtils.defaultString(matcher.group(1));
			if (!contentStr.startsWith(specStr)) {
				return specStr + contentStr;
			}
			return contentStr;
		}

		return StringUtils.defaultString(content);
	}
}

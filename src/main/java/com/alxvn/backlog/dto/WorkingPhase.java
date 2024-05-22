/**
 *
 */
package com.alxvn.backlog.dto;

import java.util.Collections;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import com.fasterxml.jackson.annotation.JsonValue;

/**
 * @author KEDD
 *
 */
public enum WorkingPhase implements EnumCodeable {
	ID6("6", "Basic Design") //
	, ID8("8", "Outline Design") //
	, ID9("9", "Detail Design") //
	, ID10("10", "PG") // PG
	, ID11("11", "UT") //
	, ID24("24", "Delivery") // Delivery
	, ID40("40", "QA") //
	, ID41("41", "Acceptance Verification") //
	, ID42("42", "No Class") //
	, ID43("43", "Bug") // Bug
	, ID44("44", "Translate") //
	, ID45("45", "Spec") //
	, ID46("46", "Research") //
	, ID47("47", "Estimate") //
	, ID48("48", "Review") //
	, ID49("49", "Test") //
	, ID50("50", "Progress management") //
	, ID51("51", "Progress report") //
	, ID52("52", "Internal meeting") //
	, ID53("53", "Meeting with customer") //
	, ID54("54", "Environment") // Setup environment
	, ID55("55", "Problem report") //
	, ID56("56", "Learning") //
	, ID57("57", "Internal member support") //
	, ID58("58", "Outsource support") //
	, ID0("0", "NaN") //

	/* */
	;

	private static final Map<String, WorkingPhase> stringToEnum;

	static {
		final Map<String, WorkingPhase> m = new ConcurrentHashMap<>();
		for (final WorkingPhase entry : WorkingPhase.values()) {
			m.put(entry.code, entry);
		}
		stringToEnum = Collections.unmodifiableMap(m);
	}

	public static WorkingPhase fromString(final String s) {
		return stringToEnum.get(s);
	}

	private final String code;
	private final String name;

	WorkingPhase(final String code, final String name) {
		this.code = code;
		this.name = name;
	}

	@JsonValue
	@Override
	public String getCode() {
		return code;
	}

	public String getName() {
		return name;
	}

}

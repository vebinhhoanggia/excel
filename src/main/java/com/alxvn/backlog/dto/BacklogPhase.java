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
public enum BacklogPhase implements EnumCodeable {
	P_1("タスク") //
	, P_2("要件定義") //
	, P_3("設計") //
	, P_4("開発") //
	, P_5("開発(課題)") //
	, P_6("リリース") //
	, P_7("インストール＆設定") //
	, P_8("調査") //
	, P_9("調整") //
	, P_10("ドキュメント") //
	, P_11("Q&A") //
	, P_12("翻訳") //
	, P_13("JP&VN定例会") //
	, P_14("イベント") //
	/* */
	;

	private static final Map<String, BacklogPhase> stringToEnum;

	static {
		final Map<String, BacklogPhase> m = new ConcurrentHashMap<>();
		for (final BacklogPhase entry : BacklogPhase.values()) {
			m.put(entry.code, entry);
		}
		stringToEnum = Collections.unmodifiableMap(m);
	}

	public static BacklogPhase fromString(final String s) {
		return stringToEnum.get(s);
	}

	private final String code;

	BacklogPhase(String code) {
		this.code = code;
	}

	@JsonValue
	@Override
	public String getCode() {
		return code;
	}

}

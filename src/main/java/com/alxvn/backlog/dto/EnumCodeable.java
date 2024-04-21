/**
 *
 */
package com.alxvn.backlog.dto;

/**
 * @author KEDD
 *
 */
public interface EnumCodeable {

	String getCode();

	default boolean isSameCode(final String code) {
		return getCode() != null && getCode().equals(code);
	}
}

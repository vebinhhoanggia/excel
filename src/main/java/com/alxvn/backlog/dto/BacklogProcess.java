/**
 *
 */
package com.alxvn.backlog.dto;

import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;

/**
 *
 */
public class BacklogProcess {

	private String code;
	private String name;

	public static BacklogProcess of(BacklogPhase e) {
		return new BacklogProcess(e);
	}

	public static BacklogProcess of(String code, String name) {
		final var rs = new BacklogProcess();
		rs.code = code;
		rs.name = name;
		return rs;
	}

	public BacklogProcess(BacklogPhase e) {
		code = e.getCode();
		name = e.getCode();
	}

	public BacklogProcess() {
	}

	public final String getCode() {
		return code;
	}

	public final void setCode(String code) {
		this.code = code;
	}

	public final String getName() {
		return name;
	}

	public final void setName(String name) {
		this.name = name;
	}

	@Override
	public int hashCode() {
		return HashCodeBuilder.reflectionHashCode(this);
	}

	@Override
	public boolean equals(final Object obj) {
		return EqualsBuilder.reflectionEquals(this, obj);
	}

	@Override
	public String toString() {
		return ToStringBuilder.reflectionToString(this, ToStringStyle.SHORT_PREFIX_STYLE);
	}

}

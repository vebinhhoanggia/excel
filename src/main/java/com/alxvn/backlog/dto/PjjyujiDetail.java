/**
 *
 */
package com.alxvn.backlog.dto;

import java.time.LocalDate;

import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;

/**
 * @author KEDD
 *
 */
public class PjjyujiDetail {
	private String id;
	private String name;
	private String mailId;
	private LocalDate targetYmd;
	private String pjCd;
	private String pjCdJp;
	private String pjName;
	private WorkingProcess process;
	private String content;
	private String ankenNo;
	private Integer minute;

	public static class Builder {
		private String id;
		private String name;
		private String mailId;
		private LocalDate targetYmd;
		private String pjCd;
		private String pjCdJp;
		private String pjName;
		private WorkingProcess process;
		private String content;
		private String ankenNo;
		private Integer minute;

		public final Builder setMailId(String mailId) {
			this.mailId = mailId;
			return this;
		}

		public final Builder setName(String name) {
			this.name = name;
			return this;
		}

		public final Builder setId(String id) {
			this.id = id;
			return this;
		}

		public final Builder setTargetYmd(LocalDate targetYmd) {
			this.targetYmd = targetYmd;
			return this;
		}

		public final Builder setPjCd(String pjCd) {
			this.pjCd = pjCd;
			return this;
		}

		public final Builder setPjCdJp(String pjCdJp) {
			this.pjCdJp = pjCdJp;
			return this;
		}

		public final Builder setPjName(String pjName) {
			this.pjName = pjName;
			return this;
		}

		public final Builder setProcess(WorkingProcess process) {
			this.process = process;
			return this;
		}

		public final Builder setContent(String content) {
			this.content = content;
			return this;
		}

		public final Builder setAnkenNo(String ankenNo) {
			this.ankenNo = ankenNo;
			return this;
		}

		public final Builder setMinute(Integer minute) {
			this.minute = minute;
			return this;
		}

		public PjjyujiDetail build() {
			return new PjjyujiDetail(id, name, mailId, targetYmd, pjCd, pjCdJp, pjName, process, content, ankenNo,
					minute);
		}
	}

	private PjjyujiDetail(final String id, final String name, final String mailId, final LocalDate targetYmd,
			final String pjCd, final String pjCdJp, final String pjName, final WorkingProcess process,
			final String content, final String ankenNo, final Integer minute) {
		this.id = id;
		this.name = name;
		this.mailId = mailId;
		this.targetYmd = targetYmd;
		this.pjCd = pjCd;
		this.pjCdJp = pjCdJp;
		this.pjName = pjName;
		this.process = process;
		this.content = content;
		this.ankenNo = ankenNo;
		this.minute = minute;
	}

	public final String getId() {
		return id;
	}

	public final void setId(String id) {
		this.id = id;
	}

	public final String getName() {
		return name;
	}

	public final void setName(String name) {
		this.name = name;
	}

	public final String getMailId() {
		return mailId;
	}

	public final void setMailId(String mailId) {
		this.mailId = mailId;
	}

	public final LocalDate getTargetYmd() {
		return targetYmd;
	}

	public final void setTargetYmd(LocalDate targetYmd) {
		this.targetYmd = targetYmd;
	}

	public final String getPjCd() {
		return pjCd;
	}

	public final void setPjCd(String pjCd) {
		this.pjCd = pjCd;
	}

	public final String getPjCdJp() {
		return pjCdJp;
	}

	public final void setPjCdJp(String pjCdJp) {
		this.pjCdJp = pjCdJp;
	}

	public final String getPjName() {
		return pjName;
	}

	public final void setPjName(String pjName) {
		this.pjName = pjName;
	}

	public final WorkingProcess getProcess() {
		return process;
	}

	public final void setProcess(WorkingProcess processCd) {
		process = processCd;
	}

	public final String getContent() {
		return content;
	}

	public final void setContent(String content) {
		this.content = content;
	}

	public final String getAnkenNo() {
		return ankenNo;
	}

	public final void setAnkenNo(String ankenNo) {
		this.ankenNo = ankenNo;
	}

	public final Integer getMinute() {
		return minute;
	}

	public final void setMinute(Integer minute) {
		this.minute = minute;
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

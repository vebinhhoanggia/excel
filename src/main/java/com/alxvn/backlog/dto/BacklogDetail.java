/**
 *
 */
package com.alxvn.backlog.dto;

import java.math.BigDecimal;
import java.time.LocalDate;

import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;

/**
 *
 */
public class BacklogDetail {

	private String key;
	private String ankenNo;
	private String mailId;
	private String pjCdJp;
	private String issueType;
	private String subject;
	private String category;
	private String version;
	private String milestone;
	private String assignee;
	private String parentKey;
	private LocalDate expectedStartDate;
	private LocalDate expectedDueDate;
	private LocalDate actualStartDate;
	private LocalDate actualDueDate;
	private BigDecimal estimatedHours;
	private BigDecimal actualHours;
	private String status;
	private String targetCustomer;
	private String progress;
	private LocalDate expectedDeliveryDate;
	private String bugCategory;
	private String bugOrigin;
	private String bugCreator;
	private String bug3rdTest;
	private String processOfWr;

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	public String getAnkenNo() {
		return ankenNo;
	}

	public void setAnkenNo(String ankenNo) {
		this.ankenNo = ankenNo;
	}

	public String getMailId() {
		return mailId;
	}

	public void setMailId(String mailId) {
		this.mailId = mailId;
	}

	public String getPjCdJp() {
		return pjCdJp;
	}

	public void setPjCdJp(String pjCdJp) {
		this.pjCdJp = pjCdJp;
	}

	public String getIssueType() {
		return issueType;
	}

	public void setIssueType(String issueType) {
		this.issueType = issueType;
	}

	public String getSubject() {
		return subject;
	}

	public void setSubject(String subject) {
		this.subject = subject;
	}

	public String getCategory() {
		return category;
	}

	public void setCategory(String category) {
		this.category = category;
	}

	public String getVersion() {
		return version;
	}

	public void setVersion(String version) {
		this.version = version;
	}

	public String getMilestone() {
		return milestone;
	}

	public void setMilestone(String milestone) {
		this.milestone = milestone;
	}

	public String getAssignee() {
		return assignee;
	}

	public void setAssignee(String assignee) {
		this.assignee = assignee;
	}

	public String getParentKey() {
		return parentKey;
	}

	public void setParentKey(String parentKey) {
		this.parentKey = parentKey;
	}

	public LocalDate getExpectedStartDate() {
		return expectedStartDate;
	}

	public void setExpectedStartDate(LocalDate expectedStartDate) {
		this.expectedStartDate = expectedStartDate;
	}

	public LocalDate getExpectedDueDate() {
		return expectedDueDate;
	}

	public void setExpectedDueDate(LocalDate expectedDueDate) {
		this.expectedDueDate = expectedDueDate;
	}

	public LocalDate getActualStartDate() {
		return actualStartDate;
	}

	public void setActualStartDate(LocalDate actualStartDate) {
		this.actualStartDate = actualStartDate;
	}

	public LocalDate getActualDueDate() {
		return actualDueDate;
	}

	public void setActualDueDate(LocalDate actualDueDate) {
		this.actualDueDate = actualDueDate;
	}

	public BigDecimal getEstimatedHours() {
		return estimatedHours;
	}

	public void setEstimatedHours(BigDecimal estimatedHours) {
		this.estimatedHours = estimatedHours;
	}

	public BigDecimal getActualHours() {
		return actualHours;
	}

	public void setActualHours(BigDecimal actualHours) {
		this.actualHours = actualHours;
	}

	public String getStatus() {
		return status;
	}

	public void setStatus(String status) {
		this.status = status;
	}

	public String getTargetCustomer() {
		return targetCustomer;
	}

	public void setTargetCustomer(String targetCustomer) {
		this.targetCustomer = targetCustomer;
	}

	public String getProgress() {
		return progress;
	}

	public void setProgress(String progress) {
		this.progress = progress;
	}

	public LocalDate getExpectedDeliveryDate() {
		return expectedDeliveryDate;
	}

	public void setExpectedDeliveryDate(LocalDate expectedDeliveryDate) {
		this.expectedDeliveryDate = expectedDeliveryDate;
	}

	public String getBugCategory() {
		return bugCategory;
	}

	public void setBugCategory(String bugCategory) {
		this.bugCategory = bugCategory;
	}

	public String getBugOrigin() {
		return bugOrigin;
	}

	public void setBugOrigin(String bugOrigin) {
		this.bugOrigin = bugOrigin;
	}

	public String getBugCreator() {
		return bugCreator;
	}

	public void setBugCreator(String bugCreator) {
		this.bugCreator = bugCreator;
	}

	public String getBug3rdTest() {
		return bug3rdTest;
	}

	public void setBug3rdTest(String bug3rdTest) {
		this.bug3rdTest = bug3rdTest;
	}

	public String getProcessOfWr() {
		return processOfWr;
	}

	public void setProcessOfWr(String processOfWr) {
		this.processOfWr = processOfWr;
	}

	public static class Builder {

		private final BacklogDetail backlogDetail;

		public Builder() {
			backlogDetail = new BacklogDetail();
		}

		public final Builder key(String key) {
			backlogDetail.key = key;
			return this;
		}

		public final Builder ankenNo(String ankenNo) {
			backlogDetail.ankenNo = ankenNo;
			return this;
		}

		public final Builder mailId(String mailId) {
			backlogDetail.mailId = mailId;
			return this;
		}

		public final Builder pjCdJp(String pjCdJp) {
			backlogDetail.pjCdJp = pjCdJp;
			return this;
		}

		public final Builder issueType(String issueType) {
			backlogDetail.issueType = issueType;
			return this;
		}

		public final Builder subject(String subject) {
			backlogDetail.subject = subject;
			return this;
		}

		public final Builder category(String category) {
			backlogDetail.category = category;
			return this;
		}

		public final Builder version(String version) {
			backlogDetail.version = version;
			return this;
		}

		public final Builder milestone(String milestone) {
			backlogDetail.milestone = milestone;
			return this;
		}

		public final Builder assignee(String assignee) {
			backlogDetail.assignee = assignee;
			return this;
		}

		public final Builder parentKey(String parentKey) {
			backlogDetail.parentKey = parentKey;
			return this;
		}

		public final Builder expectedStartDate(LocalDate expectedStartDate) {
			backlogDetail.expectedStartDate = expectedStartDate;
			return this;
		}

		public final Builder expectedDueDate(LocalDate expectedDueDate) {
			backlogDetail.expectedDueDate = expectedDueDate;
			return this;
		}

		public final Builder actualStartDate(LocalDate actualStartDate) {
			backlogDetail.actualStartDate = actualStartDate;
			return this;
		}

		public final Builder actualDueDate(LocalDate actualDueDate) {
			backlogDetail.actualDueDate = actualDueDate;
			return this;
		}

		public final Builder estimatedHours(BigDecimal estimatedHours) {
			backlogDetail.estimatedHours = estimatedHours;
			return this;
		}

		public final Builder actualHours(BigDecimal actualHours) {
			backlogDetail.actualHours = actualHours;
			return this;
		}

		public final Builder status(String status) {
			backlogDetail.status = status;
			return this;
		}

		public final Builder targetCustomer(String targetCustomer) {
			backlogDetail.targetCustomer = targetCustomer;
			return this;
		}

		public final Builder progress(String progress) {
			backlogDetail.progress = progress;
			return this;
		}

		public final Builder expectedDeliveryDate(LocalDate expectedDeliveryDate) {
			backlogDetail.expectedDeliveryDate = expectedDeliveryDate;
			return this;
		}

		public final Builder bugCategory(String bugCategory) {
			backlogDetail.bugCategory = bugCategory;
			return this;
		}

		public final Builder bugOrigin(String bugOrigin) {
			backlogDetail.bugOrigin = bugOrigin;
			return this;
		}

		public final Builder bugCreator(String bugCreator) {
			backlogDetail.bugCreator = bugCreator;
			return this;
		}

		public final Builder bug3rdTest(String bug3rdTest) {
			backlogDetail.bug3rdTest = bug3rdTest;
			return this;
		}

		public final Builder processOfWr(String processOfWr) {
			backlogDetail.processOfWr = processOfWr;
			return this;
		}

		public BacklogDetail build() {
			return backlogDetail;
		}
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

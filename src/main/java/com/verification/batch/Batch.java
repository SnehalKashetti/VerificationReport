package com.verification.batch;

import java.sql.Timestamp;

public class Batch {

	public int getBatchId() {
		return batchId;
	}

	public void setBatchId(int batchId) {
		this.batchId = batchId;
	}

	public String getBatchName() {
		return batchName;
	}

	public void setBatchName(String batchName) {
		this.batchName = batchName;
	}

	public String getBatchStatus() {
		return batchStatus;
	}

	public void setBatchStatus(String batchStatus) {
		this.batchStatus = batchStatus;
	}

	public Timestamp getEndTimestamp() {
		return endTimestamp;
	}

	public void setEndTimestamp(Timestamp endTimestamp) {
		this.endTimestamp = endTimestamp;
	}

	public Timestamp getStartTimestamp() {
		return startTimestamp;
	}

	public int getEventcount() {
		return eventcount;
	}

	public void setEventcount(int eventcount) {
		this.eventcount = eventcount;
	}

	public int getErrorcount() {
		return errorcount;
	}

	public void setErrorcount(int errorcount) {
		this.errorcount = errorcount;
	}

	public void setStartTimestamp(Timestamp startTimestamp) {
		this.startTimestamp = startTimestamp;
	}

	private int batchId;
	
	private String batchName;
	
	private String batchStatus;
	
	private Timestamp endTimestamp;
	
	private int eventcount;
	
	private int errorcount;
	
	private Timestamp startTimestamp;
}

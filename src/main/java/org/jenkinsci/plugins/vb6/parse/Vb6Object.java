package org.jenkinsci.plugins.vb6.parse;

public class Vb6Object {
	private String gUid;
	private int majorVer;
	private int minorVer;
	private int revisionVer;
	private String fileName;
	public String getgUid() {
		return gUid;
	}
	public void setgUid(String gUid) {
		this.gUid = gUid;
	}	
	public int getMajorVer() {
		return majorVer;
	}
	public void setMajorVer(int majorVer) {
		this.majorVer = majorVer;
	}
	public int getMinorVer() {
		return minorVer;
	}
	public void setMinorVer(int minorVer) {
		this.minorVer = minorVer;
	}
	public int getRevisionVer() {
		return revisionVer;
	}
	public void setRevisionVer(int revisionVer) {
		this.revisionVer = revisionVer;
	}
	public String getFileName() {
		return fileName;
	}
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}	
}

package org.jenkinsci.plugins.vb6.parse;

public class Vb6Reference {
	private String gUid;
	private int majorVer;
	private int minorVer;
	private int revisionVer;
	private String pathFileName;
	private String projectName;
	
	public String getgUid() {
		return gUid;
	}
	public void setgUid(String gUid) {
		this.gUid = gUid;
	}
	public int getMinorVer() {
		return minorVer;
	}
	public void setMinorVer(int minorVer) {
		this.minorVer = minorVer;
	}
	public int getMajorVer() {
		return majorVer;
	}
	public void setMajorVer(int majorVer) {
		this.majorVer = majorVer;
	}
	public int getRevisionVer() {
		return revisionVer;
	}
	public void setRevisionVer(int revisionVer) {
		this.revisionVer = revisionVer;
	}
	public String getPathFileName() {
		return pathFileName;
	}
	public void setPathFileName(String fileName) {
		this.pathFileName = fileName;
	}
	public String getProjectName() {
		return projectName;
	}
	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}
	public boolean checkPathFileNameExists(String pathFileName) {
		//return File.(pathFileName);
		return false;
	}
}

package org.jenkinsci.plugins.vb6.parse;
public class Vb6Class {
	private String projectName;
	private String pathFileName;
	public String getProjectName() {
		return projectName;
	}
	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}
	public String getPathFileName() {
		return pathFileName;
	}
	public void setPathFileName(String pathFileName) {
		this.pathFileName = pathFileName;
	}
	public boolean checkPathFileNameExists(String pathFileName) {
		//return File.(pathFileName);
		return false;
	}
}

package org.jenkinsci.plugins.vb6.parse;

public class Vb6Form {
	private String pathFileName;

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

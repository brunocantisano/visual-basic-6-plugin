package org.jenkinsci.plugins.vb6.DAO;

public class Copy {
	private Filterset filterset;
	private String file;
	private String toFile;
	public Filterset getFilterset() {
		return filterset;
	}
	public void setFilterset(Filterset filterset) {
		this.filterset = filterset;
	}
	public String getFile() {
		return file;
	}
	public void setFile(String file) {
		this.file = file;
	}
	public String getToFile() {
		return toFile;
	}
	public void setToFile(String toFile) {
		this.toFile = toFile;
	}
}

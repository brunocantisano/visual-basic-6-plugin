package org.jenkinsci.plugins.vb6.DAO;

public class ConfigurationPlugin {
	private String source;
	private String target;
	private String encoding;	
	private Descriptors descriptors;	
	private String preparationGoals;
	private String tagBase;
	private String releaseProfiles;
	private CheckModificationExcludes checkModificationExcludes;

	public String getSource() {
		return source;
	}
	public void setSource(String source) {
		this.source = source;
	}
	public String getTarget() {
		return target;
	}
	public void setTarget(String target) {
		this.target = target;
	}
	public String getEncoding() {
		return encoding;
	}
	public void setEncoding(String encoding) {
		this.encoding = encoding;
	}
	public Descriptors getDescriptors() {
		return descriptors;
	}
	public void setDescriptors(Descriptors descriptors) {
		this.descriptors = descriptors;
	}
	public String getPreparationGoals() {
		return preparationGoals;
	}
	public void setPreparationGoals(String preparationGoals) {
		this.preparationGoals = preparationGoals;
	}
	public String getTagBase() {
		return tagBase;
	}
	public void setTagBase(String tagBase) {
		this.tagBase = tagBase;
	}
	public String getReleaseProfiles() {
		return releaseProfiles;
	}
	public void setReleaseProfiles(String releaseProfiles) {
		this.releaseProfiles = releaseProfiles;
	}
	public CheckModificationExcludes getCheckModificationExcludes() {
		return checkModificationExcludes;
	}
	public void setCheckModificationExcludes(CheckModificationExcludes checkModificationExcludes) {
		this.checkModificationExcludes = checkModificationExcludes;
	}	
}

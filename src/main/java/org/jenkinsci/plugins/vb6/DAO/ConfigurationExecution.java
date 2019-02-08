package org.jenkinsci.plugins.vb6.DAO;

import java.util.ArrayList;

public class ConfigurationExecution {
	private Tasks tasks;
	private String includeArtifactIds;
	private String includeClassifiers;
	private String outputDirectory;
	private String overWriteReleases;
	private String overWriteSnapshots;
	private String overWriteIfNewer;		
	private String executable;	
	private String workingDirectory;	
	private ArrayList<Argument> arguments = new ArrayList<Argument>();
	
	public String getExecutable() {
		return executable;
	}
	public void setExecutable(String executable) {
		this.executable = executable;
	}
	public String getWorkingDirectory() {
		return workingDirectory;
	}
	public void setWorkingDirectory(String workingDirectory) {
		this.workingDirectory = workingDirectory;
	}
	public ArrayList<Argument> getArguments() {
		return arguments;
	}
	public void setArguments(ArrayList<Argument> arguments) {
		this.arguments = arguments;
	}
	public Tasks getTasks() {
		return tasks;
	}
	public void setTasks(Tasks tasks) {
		this.tasks = tasks;
	}
	public String getIncludeArtifactIds() {
		return includeArtifactIds;
	}
	public void setIncludeArtifactIds(String includeArtifactIds) {
		this.includeArtifactIds = includeArtifactIds;
	}
	public String getIncludeClassifiers() {
		return includeClassifiers;
	}
	public void setIncludeClassifiers(String includeClassifiers) {
		this.includeClassifiers = includeClassifiers;
	}
	public String getOutputDirectory() {
		return outputDirectory;
	}
	public void setOutputDirectory(String outputDirectory) {
		this.outputDirectory = outputDirectory;
	}
	public String getOverWriteReleases() {
		return overWriteReleases;
	}
	public void setOverWriteReleases(String overWriteReleases) {
		this.overWriteReleases = overWriteReleases;
	}
	public String getOverWriteSnapshots() {
		return overWriteSnapshots;
	}
	public void setOverWriteSnapshots(String overWriteSnapshots) {
		this.overWriteSnapshots = overWriteSnapshots;
	}
	public String getOverWriteIfNewer() {
		return overWriteIfNewer;
	}
	public void setOverWriteIfNewer(String overWriteIfNewer) {
		this.overWriteIfNewer = overWriteIfNewer;
	}
}

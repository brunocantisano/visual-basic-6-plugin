package org.jenkinsci.plugins.vb6.DAO;

public class Plugin {
	private String groupId;
	private String artifactId;
	private String version;	
	private Executions executions;
	private ConfigurationPlugin configuration;
	
	public ConfigurationPlugin getConfiguration() {
		return configuration;
	}
	public void setConfiguration(ConfigurationPlugin configuration) {
		this.configuration = configuration;
	}
	public String getGroupId() {
		return groupId;
	}
	public void setGroupId(String groupId) {
		this.groupId = groupId;
	}
	public String getArtifactId() {
		return artifactId;
	}
	public void setArtifactId(String artifactId) {
		this.artifactId = artifactId;
	}
	public String getVersion() {
		return version;
	}
	public void setVersion(String version) {
		this.version = version;
	}
	public Executions getExecutions() {
		return executions;
	}
	public void setExecutions(Executions executions) {
		this.executions = executions;
	}	
}

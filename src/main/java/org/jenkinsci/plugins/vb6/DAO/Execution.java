package org.jenkinsci.plugins.vb6.DAO;

public class Execution {
	private String id;
	private String phase;
	private Goals goals;
	private ConfigurationExecution configuration;
	
	public ConfigurationExecution getConfiguration() {
		return configuration;
	}
	public void setConfiguration(ConfigurationExecution configuration) {
		this.configuration = configuration;
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getPhase() {
		return phase;
	}
	public void setPhase(String phase) {
		this.phase = phase;
	}
	public Goals getGoals() {
		return goals;
	}
	public void setGoals(Goals goals) {
		this.goals = goals;
	}	
}

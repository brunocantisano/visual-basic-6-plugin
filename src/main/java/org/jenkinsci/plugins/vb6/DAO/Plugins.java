package org.jenkinsci.plugins.vb6.DAO;

import java.util.ArrayList;

public class Plugins {
    private ArrayList<Plugin> plugin = new ArrayList<Plugin>();

	public ArrayList<Plugin> getPlugin() {
		return plugin;
	}

	public void setPlugin(ArrayList<Plugin> plugin) {
		this.plugin = plugin;
	}   
}

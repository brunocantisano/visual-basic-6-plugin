package org.jenkinsci.plugins.vb6.DAO;

import java.util.ArrayList;

public class Filterset {
	private ArrayList<Filter> filter = new ArrayList<Filter>();

	public ArrayList<Filter> getFilter() {
		return filter;
	}

	public void setFilter(ArrayList<Filter> filter) {
		this.filter = filter;
	}	
}

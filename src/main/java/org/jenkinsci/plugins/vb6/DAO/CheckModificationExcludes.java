package org.jenkinsci.plugins.vb6.DAO;

import java.util.ArrayList;

public class CheckModificationExcludes {
	private ArrayList<String> checkModificationExclude = new ArrayList<String>();

	public ArrayList<String> getCheckModificationExclude() {
		return checkModificationExclude;
	}

	public void setCheckModificationExclude(ArrayList<String> checkModificationExclude) {
		this.checkModificationExclude = checkModificationExclude;
	}	
}

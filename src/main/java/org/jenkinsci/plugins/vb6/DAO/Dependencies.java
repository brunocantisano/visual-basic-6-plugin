package org.jenkinsci.plugins.vb6.DAO;

import java.util.ArrayList;

public class Dependencies {
    private ArrayList<Dependency> dependency = new ArrayList<Dependency>();
   
    public ArrayList<Dependency> getDependency ()
    {
        return dependency;
    }
    
    public void setDependency (ArrayList<Dependency> dependency)
    {
        this.dependency = dependency;
    }
}

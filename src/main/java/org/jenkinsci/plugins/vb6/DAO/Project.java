package org.jenkinsci.plugins.vb6.DAO;

public class Project
{
    private Dependencies dependencies;
    private String xmlns;
    private String schemaLocation;
    private String version;
    
    //nao sao usados. Sao mantidos para fazer o parser
	private String modelVersion;
    private String groupId;
    private String artifactId;
    private String name;
    private String packaging;
    private Developers developers;
    private Scm scm;
	private DistributionManagement distributionManagement;
	private Build build;
	//////////////////////////////////////////////////
	
    public Dependencies getDependencies ()
    {
        return dependencies;
    }
    
    public void setDependencies (Dependencies dependencies)
    {
        this.dependencies = dependencies;
    }
    
    public String getXmlns ()
    {
        return xmlns;
    }

    public void setXmlns (String xmlns)
    {
        this.xmlns = xmlns;
    }
    public String getSchemaLocation ()
    {
        return schemaLocation;
    }

    public void setSchemaLocation (String schemaLocation)
    {
        this.schemaLocation = schemaLocation;
    }
    public String getVersion ()
    {
        return version;
    }
    
    public void setVersion (String version)
    {
        this.version = version;
    }

    public String getModelVersion() {
		return modelVersion;
	}

	public void setModelVersion(String modelVersion) {
		this.modelVersion = modelVersion;
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

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getPackaging() {
		return packaging;
	}

	public void setPackaging(String packaging) {
		this.packaging = packaging;
	}

	public Developers getDevelopers() {
		return developers;
	}

	public void setDevelopers(Developers developers) {
		this.developers = developers;
	}

	public Scm getScm() {
		return scm;
	}

	public void setScm(Scm scm) {
		this.scm = scm;
	}

	public DistributionManagement getDistributionManagement() {
		return distributionManagement;
	}

	public void setDistributionManagement(DistributionManagement distributionManagement) {
		this.distributionManagement = distributionManagement;
	}

	public Build getBuild() {
		return build;
	}

	public void setBuild(Build build) {
		this.build = build;
	}
	
    @Override
    public String toString()
    {
        return "ClassPojo [dependencies = "+dependencies+", version = "+version+"]";
    }
}
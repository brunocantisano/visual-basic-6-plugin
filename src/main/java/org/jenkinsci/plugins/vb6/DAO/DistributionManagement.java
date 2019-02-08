package org.jenkinsci.plugins.vb6.DAO;

public class DistributionManagement {
    private Repository repository = new Repository();
    private SnapshotRepository snapshotRepository = new SnapshotRepository();
	public Repository getRepository() {
		return repository;
	}
	public void setRepository(Repository repository) {
		this.repository = repository;
	}
	public SnapshotRepository getSnapshotRepository() {
		return snapshotRepository;
	}
	public void setSnapshotRepository(SnapshotRepository snapshotRepository) {
		this.snapshotRepository = snapshotRepository;
	}    
}

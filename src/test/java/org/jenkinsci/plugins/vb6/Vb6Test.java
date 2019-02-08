package org.jenkinsci.plugins.vb6;
import static org.junit.Assert.*;
import java.io.IOException;
import org.junit.Test;
import org.jenkinsci.plugins.vb6.parse.VbpParse;
import org.junit.Before;
public class Vb6Test {

	@Before 
	public void initialize() {		

    }
	
	@Test
	public void testReadReferences() throws IOException {		
		String path = "D:/github/visual-basic-6-plugin/src/test/resources/vb6/applications/Copia/Copia.vbp";
		VbpParse.parse(path);
		assertTrue(true);
	}
	//@Test
	public void testReadObjects() throws IOException {	
		//assertTrue();	
	}
	//@Test
	public void testReadForms() throws IOException {	
		//assertTrue();	
	}
	//@Test
	public void testReadModules() throws IOException {	
		//assertTrue();	
	}
	//@Test
	public void testReadClasses() throws IOException {	
		//assertTrue();	
	}
	//@Test
	public void testReadMajorVer() throws IOException {	
		//assertTrue();	
	}
	//@Test
	public void testReadMinorVer() throws IOException {	
		//assertTrue();	
	}
	//@Test
	public void testReadRevisionVer() throws IOException {	
		//assertTrue();	
	}
}
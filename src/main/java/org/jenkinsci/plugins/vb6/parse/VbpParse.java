package org.jenkinsci.plugins.vb6.parse;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class VbpParse {
	public static final String FILE_NOT_FOUND = "File not found!";
	public static final String COULD_NOT_READ_LINE = "Could not read line";
	public static final String COULD_NOT_CLOSE_FILE = "Could not close file";

	private static ParsedElement findPattern(String pattern, String line) {
		ParsedElement parsedObject = new ParsedElement();
		List<String> elements = new ArrayList<String>();
		Pattern patt = Pattern.compile("^(Reference=)(\\*\\\\G{.*})#([0-9]{1,4})\\.([0-9]{1,4})#([0-9]{1,4})#(.*)#(.*)$");
		Matcher match = patt.matcher(line);
		if (match != null){
			if(match.find()) {            	
				for(int index = 0; index < match.groupCount(); index++) {
					elements.add(match.group(index));
				}            
			}            
		}
		return parsedObject;
	}
	public static String findType(String line) {
		String pattern = "^(Type)=(.*)$";
		ParsedElement parsed = findPattern(pattern, line);		
		return parsed.getElement().get(2);
	}
	public static Vb6Reference findReference(String line) {
		//^(Reference)=(\*\\G{.*})#([0-9]{1,4})\.([0-9]{1,4})#([0-9]{1,4})#(.*)#(.*)$
		Vb6Reference reference = new Vb6Reference();
		String pattern = "^(Reference)=(\\*\\\\G{.*})#([0-9]{1,4})\\.([0-9]{1,4})#([0-9]{1,4})#(.*)#(.*)$";
		ParsedElement parsed = findPattern(pattern, line);		
		reference.setgUid(parsed.getElement().get(2));
		reference.setMajorVer(Integer.parseInt(parsed.getElement().get(3)));
		reference.setMinorVer(Integer.parseInt(parsed.getElement().get(4)));
		reference.setRevisionVer(Integer.parseInt(parsed.getElement().get(5)));
		reference.setPathFileName(parsed.getElement().get(6));
		reference.setProjectName(parsed.getElement().get(7));    	
		return reference;
	}
	public static Vb6Object findObject(String line) {
		//^(Object)=(\{.*\})#([0-9]{1,4})\.([0-9]{1,4})#([0-9]{1,4})\;(\s.*)$		
		Vb6Object obj = new Vb6Object();
		String pattern = "^(Object)=(\\{.*\\})#([0-9]{1,4})\\.([0-9]{1,4})#([0-9]{1,4})\\;(\\s.*)$";
		ParsedElement parsed = findPattern(pattern, line);		
		obj.setgUid(parsed.getElement().get(2));
		obj.setMajorVer(Integer.parseInt(parsed.getElement().get(3)));
		obj.setMinorVer(Integer.parseInt(parsed.getElement().get(4)));
		obj.setRevisionVer(Integer.parseInt(parsed.getElement().get(5)));
		obj.setFileName(parsed.getElement().get(6));
		return obj;
	}
	public static Vb6Form findForm(String line) {
		//^(Form)=(.*)$
		Vb6Form frm = new Vb6Form();
		String pattern = "^(Form)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);		
		frm.setPathFileName(parsed.getElement().get(2));
		return frm;
	}
	public static Vb6Module findModule(String line) {
		//^(Module)=(.*)\;(.*)$
		Vb6Module mdl = new Vb6Module();
		String pattern = "^(Module)=(.*)\\;(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);		
		mdl.setProjectName(parsed.getElement().get(2));
		mdl.setPathFileName(parsed.getElement().get(3));
		return mdl;
	}
	public static Vb6Class findClass(String line) {
		//^(Class)=(.*)\;(.*)$
		Vb6Class cls = new Vb6Class();
		String pattern = "^(Class)=(.*)\\;(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);		
		cls.setProjectName(parsed.getElement().get(2));
		cls.setPathFileName(parsed.getElement().get(3));
		return cls;
	}
	public static int findMajorVer(String line) {
		//^(MajorVer)=(.*)$
		String pattern = "^(MajorVer)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return Integer.parseInt(parsed.getElement().get(2));
	}
	public static int findMinorVer(String line) {
		//^(MinorVer)=(.*)$
		String pattern = "^(MinorVer)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return Integer.parseInt(parsed.getElement().get(2));
	}
	public static int findRevisionVer(String line) {
		//^(RevisionVer)=(.*)$
		String pattern = "^(RevisionVer)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return Integer.parseInt(parsed.getElement().get(2));
	}
	public static int findAutoIncrementer(String line) {
		//^(AutoIncrementVer)=(.*)$
		String pattern = "^(AutoIncrementVer)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return Integer.parseInt(parsed.getElement().get(2));
	}
	public static String findVersionCompanyName(String line) {
		//^(VersionCompanyName)=(.*)$
		String pattern = "^(VersionCompanyName)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findStartUp(String line) {
		//^(Startup)=(.*)$
		String pattern = "^(Startup)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findHelpFile(String line) {
		//^(HelpFile)=(.*)$
		String pattern = "^(HelpFile)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findTitle(String line) {
		//^(Title)=(.*)$
		String pattern = "^(Title)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findExeName32(String line) {
		//^(ExeName32)=(.*)$
		String pattern = "^(ExeName32)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findPath32(String line) {
		//^(Path32)=(.*)$
		String pattern = "^(Path32)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findCommand32(String line) {
		//^(Command32)=(.*)$
		String pattern = "^(Command32)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static String findName(String line) {
		//^(Name)=(.*)$
		String pattern = "^(Name)=(.*)$";		
		ParsedElement parsed = findPattern(pattern, line);
		return parsed.getElement().get(2);
	}
	public static boolean parse(String vbpPathFileName) {
		BufferedReader br = null;
		try {
			br = new BufferedReader(new FileReader(vbpPathFileName));
		} catch (FileNotFoundException e1) {
			System.out.println(FILE_NOT_FOUND);
		}
		try {
			String line = "";
			try {
				line = br.readLine();
				while (line != null) {
					findType(line);
					findReference(line);
					findObject(line);
					findForm(line);
					findModule(line);
					findClass(line);
					findMajorVer(line);
					findMinorVer(line);
					findRevisionVer(line);
					findAutoIncrementer(line);
					findVersionCompanyName(line);
					findStartUp(line);
					findHelpFile(line);
					findTitle(line);
					findExeName32(line);
					findPath32(line);
					findCommand32(line);
					findName(line);
					line = br.readLine();					
				}
			}
			catch (IOException e) {
				System.out.println(COULD_NOT_READ_LINE);
			}		    
		} 
		finally {
			try {
				br.close();
				return true;
			} catch (IOException e) {
				System.out.println(COULD_NOT_CLOSE_FILE);
			}
		}
		return false;
	}
}

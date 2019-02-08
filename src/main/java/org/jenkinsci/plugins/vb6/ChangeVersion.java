package org.jenkinsci.plugins.vb6;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.Reader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import org.apache.commons.io.FilenameUtils;
import org.jenkinsci.plugins.vb6.DAO.Dependencies;
import org.jenkinsci.plugins.vb6.DAO.Project;
import hudson.FilePath;

public class ChangeVersion {
	private File TempFile;
	private Project project = null;
	private String caminhoVbp = null;
	private String fileNamePom = null;	
	private char[] majorVer = new char[4];
    private char[] minorVer = new char[4];
    private char[] revisionVer = new char[4];
    private FilePath workspace = null;
    
	public ChangeVersion(FilePath workspace){
		this.majorVer = "0000".toCharArray();
    	this.minorVer = "0000".toCharArray();
        this.revisionVer = "0000".toCharArray();
        this.workspace = workspace;
    }
	   
	public boolean isNumeric(String s) {  
	    return s.matches("^[0-9]{1,4}$");  
	}  
	public String getCaminhoVbp() {
		return caminhoVbp;
	}
	public String getFileNamePom() {
		return fileNamePom;
	}
    public File getTempFile() {
		return TempFile;
	}
	public String getMajorVer() {
		return majorVer.toString();
	}
    /**
     * Se numero maior que 9999, sera truncado
     * @param majorVer parte mais significativa da versao
     * @return true se conseguiu atribuir / false se nao conseguiu atribuir
     */
	public boolean setMajorVer(char[] majorVer) {
		//checa se nao e numero de 0 a 9
		if (!isNumeric(majorVer.toString())) return false;			
		this.majorVer = majorVer;
		return true;
	}
	public String getMinorVer() {
		return minorVer.toString();
	}
	public boolean setMinorVer(char[] minorVer) {
		//checa se nao e numero de 0 a 9
		if (!isNumeric(minorVer.toString())) return false;			
		this.minorVer = minorVer;
		return true;
	}
	public String getRevisionVer() {
		return revisionVer.toString();
	}
	public boolean setRevisionVer(char[] revisionVer) {
		//checa se nao e numero de 0 a 9
		if (!isNumeric(revisionVer.toString())) return false;			
		this.revisionVer = revisionVer;
		return true;
	}
	
	public boolean wasVbpChanged(){
	    StringBuffer strVbp = new StringBuffer();
	    int pos = 0;
	    int posFim = 0;    
		String line = "";	
		
		try{	
		    if (!this.workspace.exists()){
		        //O path do vbp nao existe
		        return false;
		    }	    	    
		     
		    //cria arquivo temporario
		    TempFile = File.createTempFile("tmp", FilenameUtils.getExtension(this.workspace.getRemote()), new File(this.workspace.getRemote()));
		    
		    InputStream inputStream       = new FileInputStream(caminhoVbp);
		    Reader      inputStreamReader = new InputStreamReader(inputStream, "UTF-8");
		    BufferedReader reader = new BufferedReader(inputStreamReader);
			while ((line = reader.readLine()) != null)
			{
		        //Altera os campos de acordo com a versao que veio da TAG_VERSION
		        pos = line.indexOf("MajorVer");
		        if (pos > 0) {
		            posFim = line.lastIndexOf("=");
		            line = line.substring(1, posFim) + majorVer.toString();
			    }
		        pos = line.indexOf("MinorVer");
		        if (pos > 0) {
		            posFim = line.lastIndexOf("=");
		            line = line.substring(1, posFim) + majorVer.toString();
			    }
		        pos = line.indexOf("RevisionVer");
		        if (pos > 0){
		            posFim = line.lastIndexOf("=");
		            line = line.substring(1, posFim) + revisionVer.toString();
		        }
		        strVbp.append(line).append('\n');
			}
			inputStreamReader.close();
			reader.close();	
						
		    //sobrescreve o arquivo vbp com o conteudo de strVbp
			OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(caminhoVbp), "UTF-8");
			String contentInBytes = strVbp.toString().substring(0, strVbp.length());
			PrintWriter pw = new PrintWriter(osw);
			pw.println(contentInBytes);
			pw.close();			      
		}catch(Exception e){
			System.err.println(e);
			return false;
		}
		return true;
	}
	public boolean getParserVersionTag(){
	    int posSnaphot = 0;
	    String tag_snap = "";
	    DateFormat df = new SimpleDateFormat("mmss");
	    Date today = Calendar.getInstance().getTime();        
	    String timestamp = df.format(today);
	    String arrVersion[];
		try{
		    tag_snap = "-SNAPSHOT";
		    arrVersion = project.getVersion().split("\\.");
		    
	        majorVer =arrVersion[0].toCharArray();
            if (arrVersion[1].length() == 0){
                minorVer = "0000".toCharArray();
            }else{
            	posSnaphot = arrVersion[1].indexOf(tag_snap);
            	if(posSnaphot > 0){
            		minorVer = arrVersion[1].replace(tag_snap, "").toCharArray();
            		revisionVer = timestamp.toCharArray();
            	}else{
            		minorVer = arrVersion[1].toCharArray();
            	}
            }
            if (arrVersion[2].length() == 0){
                revisionVer = "0000".toCharArray();
            }else{
            	posSnaphot = arrVersion[2].indexOf(tag_snap);
            	if(posSnaphot > 0){
            		revisionVer = timestamp.toCharArray();
            	}else{
            		revisionVer = arrVersion[2].toCharArray();
            	}
            }		    
		}catch(Exception e){
			return false;
		}
		return true;
	}	
	public String getVersionFromPom(){
		return project.getVersion();
	}
	public Dependencies getDependenciesFromPom(){
		return project.getDependencies(); 
	}
}

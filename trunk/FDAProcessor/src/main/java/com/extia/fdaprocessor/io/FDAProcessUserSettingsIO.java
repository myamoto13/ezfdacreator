package com.extia.fdaprocessor.io;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.log4j.Logger;

import com.extia.fdaprocessor.data.FDAProcessUserSettings;

public class FDAProcessUserSettingsIO {
	
	static Logger logger = Logger.getLogger(FDAProcessUserSettingsIO.class);
	
	private String configFilePath;
	
	private String getConfigFilePath() {
		return configFilePath;
	}

	public void setConfigFilePath(String configFilePath) {
		this.configFilePath = configFilePath;
	}

	public FDAProcessUserSettings readScrappingSettings() throws IOException {
		FDAProcessUserSettings result = null;
		String configFilePath = getConfigFilePath();
		if(configFilePath != null){
			result = new FDAProcessUserSettings();
			Properties prop = new Properties();
			FileInputStream fs = null;
			try {
	    		fs = new FileInputStream(new File(configFilePath));
	    		
	    		prop.load(fs);
	    		
	    		String srcFilePath = prop.getProperty("srcFilePath");
	    		if(srcFilePath != null){
	    			result.setSrcDir(srcFilePath);
	    		}else{
	    			logger.warn("srcFilePath not found.");
	    		}
	    		
	    		String destFilePath = prop.getProperty("destFilePath");
	    		if(destFilePath != null){
	    			result.setDestDir(destFilePath);
	    		}else{
	    			logger.warn("destFilePath not found.");
	    		}
	    	} catch (IOException ex) {
	    		logger.error(ex.getMessage(), ex);
	        } finally {
	        	if(fs != null){
	        		fs.close();
	        	}
	        }
			
		}else{
			logger.warn("Config file path not set, directory settings won't be saved.");
		}
		return result;
	}
	
	public void writeScrappingSettings(FDAProcessUserSettings settings) throws IOException {
		String configFilePath = getConfigFilePath();
		if(configFilePath != null && settings != null){
			Properties prop = new Properties();
			FileOutputStream fo = null;
			try{
				fo = new FileOutputStream(new File(configFilePath));
				if(settings.getSrcDir() != null){
					prop.setProperty("srcFilePath", settings.getSrcDir());
				}
				if(settings.getDestDir() != null){
					prop.setProperty("destFilePath", settings.getDestDir());
				}
				prop.store(fo, "comments");
			} catch(IOException ex){
				throw ex;
			} finally{
				if(fo != null){
					fo.close();
				}
			}
			
		}
	}
}

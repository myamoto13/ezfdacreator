package com.extia.fdaprocessor.ui.fdaprocessor;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.extia.fdaprocessor.FDAProcessor;

public class MFDAProcessor {
	
	private List<MViadeoScraperListener> modeleListenerList;
	
	private FDAProcessor fdaProcessor;

	public MFDAProcessor() {
		modeleListenerList = new ArrayList<MFDAProcessor.MViadeoScraperListener>();
	}
	
	public void setSrcDir(File srcDir) {
		getFdaProcessor().setSrcDir(srcDir);
	}
	
	public void setDestDir(File destDir) {
		getFdaProcessor().setDestDir(destDir);
	}
	
	public FDAProcessor getFdaProcessor() {
		return fdaProcessor;
	}

	public void setFdaProcessor(FDAProcessor fdaProcessor) {
		this.fdaProcessor = fdaProcessor;
	}
	
	public void addModeleListener(MViadeoScraperListener modeleListener) {
		modeleListenerList.add(modeleListener);		
	}
	
	interface MViadeoScraperListener{
		void progressUpdated(int progress);

		void searchEnabled(boolean enabled);

	}

	public void progressUpdated(int progress) {
		for (MViadeoScraperListener modeleListener : modeleListenerList) {
			modeleListener.progressUpdated(progress);
		}
	}

	public void enableSearch(boolean enabled) {
		for (MViadeoScraperListener modeleListener : modeleListenerList) {
			modeleListener.searchEnabled(enabled);
		}
	}

	public File getDestDir() {
		return getFdaProcessor().getDestDir();
	}

	public File getSrcDir() {
		return getFdaProcessor().getSrcDir();
	}

}

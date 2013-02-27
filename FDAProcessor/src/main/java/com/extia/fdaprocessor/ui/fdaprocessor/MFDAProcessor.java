package com.extia.fdaprocessor.ui.fdaprocessor;

import java.util.ArrayList;
import java.util.List;

import com.extia.fdaprocessor.FDAProcessor;

public class MFDAProcessor {
	
	private List<MViadeoScraperListener> modeleListenerList;
	
	private FDAProcessor fdaProcessor;

	public MFDAProcessor() {
		modeleListenerList = new ArrayList<MFDAProcessor.MViadeoScraperListener>();
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
	}

}

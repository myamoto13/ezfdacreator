package com.extia.fdaprocessor.ui.fdaprocessor;

import java.util.ArrayList;
import java.util.List;

import javax.swing.SwingWorker;

import com.extia.fdaprocessor.excel.FDAProcessor;
import com.extia.fdaprocessor.excel.FDAProcessor.FDAProcessProgressListener;

class FDAProcessorTask extends SwingWorker<Void, Void> {
	
    private FDAProcessor fDAProcessor;
    private List<FDAProcessorTaskListener> FDAProcessorTaskListenerList;
	
	public FDAProcessorTask() {
		FDAProcessorTaskListenerList = new ArrayList<FDAProcessorTaskListener>();
	}
	
	public void removeFDAProcessorTaskListener(FDAProcessorTaskListener fDAProcessorTaskListener) {
		FDAProcessorTaskListenerList.remove(fDAProcessorTaskListener);
	}
	
	public void addFDAProcessorTaskListener(FDAProcessorTaskListener fDAProcessorTaskListener) {
		FDAProcessorTaskListenerList.add(fDAProcessorTaskListener);		
	}
	
    public void setFDAProcessor(FDAProcessor viadeoScraper) {
		this.fDAProcessor = viadeoScraper;
	}
    
    public Void doInBackground() {
    	
    	FDAProcessProgressListener scrapingProgressListener = new FDAProcessProgressListener(){
    		public void progressUpdated(int progress){
    			setProgress(progress);
    		}
    	};
    	
    	try{
    		setProgress(0);
    		fireStarting();
    		fDAProcessor.addFDAProcessProgressListener(scrapingProgressListener);
    		fDAProcessor.processFdas();
    	}catch(Exception ex){
    		fireError(ex);
    	}finally{
    		fDAProcessor.removeFDAProcessProgressListener(scrapingProgressListener);
    	}
    	return null;
    }
    
	public void done() {
//    	fDAProcessor.setInterruptFlag(false);
		fireFinished();
    }
    
	private void fireError(Exception ex) {
		for (FDAProcessorTaskListener scrapTaskListener : FDAProcessorTaskListenerList) {
			scrapTaskListener.error(ex);
		}		
	}
	
	private void fireFinished() {
		for (FDAProcessorTaskListener scrapTaskListener : FDAProcessorTaskListenerList) {
			scrapTaskListener.finished();
		}		
	}

	private void fireStarting() {
		for (FDAProcessorTaskListener scrapTaskListener : FDAProcessorTaskListenerList) {
			scrapTaskListener.starting();
		}		
	}
	
//	public void cancelScraping(){
//    	viadeoScraper.setInteruptFlag(true);
//    }
	
	interface FDAProcessorTaskListener{
		public void error(Exception ex);
		public void finished();
		public void starting();
	}

	
}
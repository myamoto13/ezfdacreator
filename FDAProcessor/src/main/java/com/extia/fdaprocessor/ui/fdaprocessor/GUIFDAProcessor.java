package com.extia.fdaprocessor.ui.fdaprocessor;

import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JPanel;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.extia.fdaprocessor.FDAProcessorLauncher;
import com.extia.fdaprocessor.excel.FDAProcessor;
import com.extia.fdaprocessor.excel.FDAProcessor.FDAProcessProgressListener;
import com.extia.fdaprocessor.ui.fdaprocessor.FDAProcessorTask.FDAProcessorTaskListener;
import com.extia.fdaprocessor.ui.fdaprocessor.VFDAProcessor.VFDAProcessorListener;

public class GUIFDAProcessor implements VFDAProcessorListener {
	
	static Logger logger = Logger.getLogger(FDAProcessorLauncher.class);

	private MFDAProcessor modele;
	private VFDAProcessor vue;
	
	private FDAProcessorTask fDAProcessorTask;

	private FDAProcessProgressListener fDAProcessProgressListener;
	
	private PropertyChangeListener pChangeListener;
	private FDAProcessorTaskListener fDAProcessorTaskListener;
	
	public GUIFDAProcessor(){
		MFDAProcessor modele = new MFDAProcessor();
		VFDAProcessor vue = new VFDAProcessor();

		setModele(modele);
		setVue(vue);

		vue.setModele(modele);
		vue.addVueListener(this);

		modele.addModeleListener(vue);
		
	}
	
	public void setFdaProcessor(FDAProcessor fdaProcessor) {
		getModele().setFdaProcessor(fdaProcessor);
	}

	private MFDAProcessor getModele() {
		return modele;
	}

	private void setModele(MFDAProcessor modele) {
		this.modele = modele;
	}

	private VFDAProcessor getVue() {
		return vue;
	}

	private void setVue(VFDAProcessor vue) {
		this.vue = vue;
	}

	public JPanel getUI() throws IOException{
		return getVue().getUi();
	}

	public FDAProcessorTask getfDAProcessorTask() {
		return fDAProcessorTask;
	}

	public void setfDAProcessorTask(FDAProcessorTask fDAProcessorTask) {
		this.fDAProcessorTask = fDAProcessorTask;
	}

	public void fireProcessFDA() throws InvalidFormatException, IOException {
		
		//Instances of javax.swing.SwingWorker are not reusuable
		if(getfDAProcessorTask() != null){
			getfDAProcessorTask().removePropertyChangeListener(getPChangeListener());
			getfDAProcessorTask().removeFDAProcessorTaskListener(getFDAProcessorTaskListener());
		}
		getModele().progressUpdated(0);
		FDAProcessorTask fDAProcessorTask = new FDAProcessorTask();
		setfDAProcessorTask(fDAProcessorTask);
		fDAProcessorTask.setFDAProcessor(getModele().getFdaProcessor());
		fDAProcessorTask.addPropertyChangeListener(getPChangeListener());
		getfDAProcessorTask().addFDAProcessorTaskListener(getFDAProcessorTaskListener());
		fDAProcessorTask.execute();
	}

	private FDAProcessorTaskListener getFDAProcessorTaskListener() {
		if(fDAProcessorTaskListener == null){
			fDAProcessorTaskListener = new FDAProcessorTaskListener(){

				public void error(Exception ex) {
					logger.error(ex.getMessage(), ex);
				}

				public void finished() {
					getModele().enableSearch(true);
				}

				public void starting() {
					getModele().enableSearch(false);
				}
				
			};
		}
		
		return fDAProcessorTaskListener;
	}

	private PropertyChangeListener getPChangeListener(){
		if(pChangeListener == null){
			pChangeListener = new PropertyChangeListener() {
				public void propertyChange(PropertyChangeEvent evt) {
					if ("progress" == evt.getPropertyName()) {
						int progress = (Integer) evt.getNewValue();
						getModele().progressUpdated(progress);
					}
				}

			};
		}
		return pChangeListener;
	}

	public void fireDestDirUpdated(File destDir) {
		getModele().setDestDir(destDir);
	}

	public void fireSrcDirUpdated(File srcDir) {
		getModele().setSrcDir(srcDir);
	}

}

package com.extia.fdaprocessor.ui.fdaprocessor;

import java.io.File;
import java.io.IOException;

import javax.swing.JPanel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.extia.fdaprocessor.FDAProcessor;
import com.extia.fdaprocessor.ui.fdaprocessor.VFDAProcessor.VFDAProcessorListener;

public class GUIFDAProcessor implements VFDAProcessorListener {

	private MFDAProcessor modele;
	private VFDAProcessor vue;
	
	private FDAProcessor fdaProcessor;
	
	public GUIFDAProcessor(){
		MFDAProcessor modele = new MFDAProcessor();
		VFDAProcessor vue = new VFDAProcessor();

		setModele(modele);
		setVue(vue);

		vue.setModele(modele);
		vue.addVueListener(this);

		modele.addModeleListener(vue);
		
		fdaProcessor = new FDAProcessor();
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

	public void fireScrapingStopped() {
		
	}

	public void fireProcessFDA(File srcDir, File destDir) throws InvalidFormatException, IOException {
		fdaProcessor.processFdas(srcDir, destDir);
	}


}

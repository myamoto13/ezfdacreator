package com.extia.fdaprocessor.ui.fdaprocessor;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.ObjectInputStream.GetField;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.extia.fdaprocessor.ui.fdaprocessor.MFDAProcessor.MViadeoScraperListener;

public class VFDAProcessor implements MViadeoScraperListener {

	private MFDAProcessor modele;

	private JPanel ui;

	private JButton srcDirChooseBtn;
	private JFileChooser srcDirFileChooser;

	private JButton destDirChooseBtn;
	private JFileChooser destDirFileChooser;

	private JButton processFDABtn;
	
	private List<VFDAProcessorListener> vueListenerList;

	private JTextField srcDirTxtFld;
	private JTextField destDirTxtFld;

	public VFDAProcessor() {
		vueListenerList = new ArrayList<VFDAProcessorListener>();
	}

	public void addVueListener(VFDAProcessorListener vueListener) {
		vueListenerList.add(vueListener);		
	}

	public void setModele(MFDAProcessor modele) {
		this.modele = modele;		
	}

	private MFDAProcessor getModele() {
		return modele;
	}

	private ActionListener getActionListener(){
		return new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try{
					if (e.getSource() == srcDirChooseBtn) {
						JFileChooser fc = getSrcDirFileChooser();

						int returnVal = fc.showOpenDialog(getUi());
						if (returnVal == JFileChooser.APPROVE_OPTION) {
							srcDirTxtFld.setText(fc.getSelectedFile().getAbsolutePath());
						}
					}else if(e.getSource() == destDirChooseBtn){
						JFileChooser fc = getDestDirFileChooser();

						int returnVal = fc.showOpenDialog(getUi());
						if (returnVal == JFileChooser.APPROVE_OPTION) {
							destDirTxtFld.setText(fc.getSelectedFile().getAbsolutePath());
						}
					}else if(e.getSource() == processFDABtn){
						fireProcessFDA();
					}
				}catch(Exception ex){
					ex.printStackTrace();
				}
			}
		};
	}
	
	private JFileChooser getDestDirFileChooser() {
		if(destDirFileChooser == null){
			destDirFileChooser = new JFileChooser();
			destDirFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		}
		return destDirFileChooser;
	}

	private JFileChooser getSrcDirFileChooser() {
		if(srcDirFileChooser == null){
			srcDirFileChooser = new JFileChooser();
			srcDirFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		}
		return srcDirFileChooser;
	}

	private FileFilter getDirectoryFilter() {
		return  new FileFilter() {

			public String getDescription() {
				return "répertoire";
			}

			public boolean accept(File file) {
				return file != null && file.isDirectory();
			}
		};
	}

	public JPanel getUi() throws IOException {
		if(ui == null){
			ActionListener al = getActionListener();
			
			processFDABtn = new JButton("Générer FDA");
			processFDABtn.addActionListener(al);
			
			srcDirChooseBtn = new JButton("Choisir source");
			srcDirChooseBtn.addActionListener(al);

			destDirChooseBtn = new JButton("Choisir destination");
			destDirChooseBtn.addActionListener(al);

			srcDirTxtFld = new JTextField();
			srcDirTxtFld.setPreferredSize(new Dimension(150, 30));
			srcDirTxtFld.setEnabled(false);
			
			destDirTxtFld = new JTextField();
			destDirTxtFld.setPreferredSize(srcDirTxtFld.getPreferredSize());
			destDirTxtFld.setEnabled(false);

			JPanel pnlBtnProcess = new JPanel(new GridBagLayout());
			pnlBtnProcess.setOpaque(false);
			pnlBtnProcess.add(processFDABtn, new GridBagConstraints(0, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));
			
			ui = new JPanel(new GridBagLayout());
			ui.setOpaque(false);
			
			ui.add(pnlBtnProcess, new GridBagConstraints(0, 0, 2, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.HORIZONTAL, new Insets(0, 0, 0, 0), 0, 0));
			
			ui.add(srcDirChooseBtn, new GridBagConstraints(0, 1, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));
			ui.add(srcDirTxtFld, new GridBagConstraints(1, 1, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));
			
			ui.add(destDirChooseBtn, new GridBagConstraints(0, 2, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));
			ui.add(destDirTxtFld, new GridBagConstraints(1, 2, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));
			
		}
		return ui;
	}

	private JFrame getParentFrame() {
		Component p = ui;
		while ( (p = p.getParent()) != null && !(p instanceof JFrame));
		return((JFrame)p);
	}

	interface VFDAProcessorListener{
		public void fireProcessFDA(File srcDir, File destDir) throws InvalidFormatException, IOException;
		public void fireScrapingStopped();
	}
	
	private void fireProcessFDA() throws InvalidFormatException, IOException {
		for (VFDAProcessorListener vueListener : vueListenerList) {
			vueListener.fireProcessFDA(getSrcDirFileChooser().getSelectedFile(), getDestDirFileChooser().getSelectedFile());
		}
	}

}

package com.extia.fdaprocessor.ui.fdaprocessor;

import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JTextField;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.extia.fdaprocessor.ui.fdaprocessor.MFDAProcessor.MViadeoScraperListener;

public class VFDAProcessor implements MViadeoScraperListener {

	private MFDAProcessor modele;

	private JPanel ui;

	private JButton srcDirChooseBtn;
	private JFileChooser srcDirFileChooser;

	private JButton destDirChooseBtn;
	private JFileChooser destDirFileChooser;

	private JProgressBar progressBar;
	
	private JButton processFDABtn;
	private JButton makeSyntheseBtn;
	
	private List<VFDAProcessorListener> vueListenerList;

	private JTextField srcDirTxtFld;
	private JTextField destDirTxtFld;

	private JPanel progressPnl;
	
	private JPanel ctrlPnl;
	
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
							fireSrcDirUpdated(fc.getSelectedFile());
						}
					}else if(e.getSource() == destDirChooseBtn){
						JFileChooser fc = getDestDirFileChooser();

						int returnVal = fc.showOpenDialog(getUi());
						if (returnVal == JFileChooser.APPROVE_OPTION) {
							destDirTxtFld.setText(fc.getSelectedFile().getAbsolutePath());
							fireDestDirUpdated(fc.getSelectedFile());
						}
					}else if(e.getSource() == processFDABtn){
						fireProcessFDA();
					}else if(e.getSource() == makeSyntheseBtn){
						fireMakeSyntheseFDA();
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

	public JPanel getUi() throws IOException {
		if(ui == null){
			ActionListener al = getActionListener();
			
			processFDABtn = new JButton(new ImageIcon(getClass().getResource("/icones/media_play_green.png")));
			processFDABtn.setToolTipText("Générer FDA");
			processFDABtn.addActionListener(al);
			
			makeSyntheseBtn = new JButton(new ImageIcon(getClass().getResource("/icones/tableau.jpg")));
			makeSyntheseBtn.setToolTipText("Faire synthèse des FDA");
			makeSyntheseBtn.addActionListener(al);
			
			
			srcDirChooseBtn = new JButton(new ImageIcon(getClass().getResource("/icones/repertoire.gif")));
			srcDirChooseBtn.setToolTipText("Choisir source");
			srcDirChooseBtn.addActionListener(al);

			destDirChooseBtn = new JButton(new ImageIcon(getClass().getResource("/icones/excel.jpg")));
			destDirChooseBtn.setToolTipText("Choisir destination");
			destDirChooseBtn.addActionListener(al);

			srcDirTxtFld = new JTextField();
			srcDirTxtFld.setPreferredSize(new Dimension(300, 30));
			srcDirTxtFld.setEnabled(false);
			srcDirTxtFld.setText(getModele().getSrcDir() != null ? getModele().getSrcDir().getAbsolutePath() : null);
			
			destDirTxtFld = new JTextField();
			destDirTxtFld.setPreferredSize(srcDirTxtFld.getPreferredSize());
			destDirTxtFld.setEnabled(false);
			destDirTxtFld.setText(getModele().getDestDir() != null ? getModele().getDestDir().getAbsolutePath() : null);
			
			JPanel cmdPnl = new JPanel(new GridBagLayout());
			cmdPnl.setOpaque(false);
			cmdPnl.add(processFDABtn, new GridBagConstraints(0, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(0, 0, 5, 0), 0, 0));
			cmdPnl.add(makeSyntheseBtn, new GridBagConstraints(1, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(0, 5, 5, 0), 0, 0));
			
			
			ctrlPnl = new JPanel(new GridBagLayout());
			ctrlPnl.setOpaque(false);
			
			ctrlPnl.add(cmdPnl, new GridBagConstraints(0, 0, 2, 1, 1, 0, GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(0, 0, 5, 0), 0, 0));
			
			ctrlPnl.add(srcDirChooseBtn, new GridBagConstraints(0, 1, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 5, 0), 0, 0));
			ctrlPnl.add(srcDirTxtFld, new GridBagConstraints(1, 1, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 5, 5, 0), 0, 0));
			
			ctrlPnl.add(destDirChooseBtn, new GridBagConstraints(0, 2, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));
			ctrlPnl.add(destDirTxtFld, new GridBagConstraints(1, 2, 1, 1, 0, 0, GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 5, 0, 0), 0, 0));
			
			progressBar = new JProgressBar(0, 100);
			progressBar.setValue(0);
			progressBar.setStringPainted(true);
			progressBar.setPreferredSize(new Dimension(ctrlPnl.getPreferredSize().width, processFDABtn.getPreferredSize().height));
			
			progressPnl = new JPanel(new GridBagLayout());
			progressPnl.setOpaque(false);
			progressPnl.setVisible(false);
			progressPnl.add(progressBar, new GridBagConstraints(1, 0, 1, 1, 1, 0, GridBagConstraints.CENTER, GridBagConstraints.HORIZONTAL, new Insets(5, 5, 0, 0), 0, 0));

			ui = new JPanel(new GridBagLayout());
			ui.setOpaque(false);
			
			ui.add(progressPnl, new GridBagConstraints(0, 0, 1, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.HORIZONTAL, new Insets(0, 0, 5, 0), 0, 0));
			ui.add(ctrlPnl, new GridBagConstraints(0, 2, 1, 1, 0, 0, GridBagConstraints.CENTER, GridBagConstraints.HORIZONTAL, new Insets(0, 0, 0, 0), 0, 0));
		}
		return ui;
	}

	interface VFDAProcessorListener{
		public void fireProcessFDA() throws InvalidFormatException, IOException;

		public void fireDestDirUpdated(File destDir);

		public void fireSrcDirUpdated(File srcDir);

		public void fireMakeSyntheseFDA();
	}
	
	
	private void fireDestDirUpdated(File destDir) {
		for (VFDAProcessorListener vueListener : vueListenerList) {
			vueListener.fireDestDirUpdated(destDir);
		}
	}
	
	private void fireMakeSyntheseFDA() {
		for (VFDAProcessorListener vueListener : vueListenerList) {
			vueListener.fireMakeSyntheseFDA();
		}
	}

	private void fireSrcDirUpdated(File srcDir) {
		for (VFDAProcessorListener vueListener : vueListenerList) {
			vueListener.fireSrcDirUpdated(srcDir);
		}
	}
	
	private void fireProcessFDA() throws InvalidFormatException, IOException {
		for (VFDAProcessorListener vueListener : vueListenerList) {
			vueListener.fireProcessFDA();
		}
	}

	public void progressUpdated(int progress) {
		progressBar.setValue(progress);
	}

	public void searchEnabled(boolean enabled) {
		progressPnl.setVisible(!enabled);
		ctrlPnl.setVisible(enabled);
	}

}

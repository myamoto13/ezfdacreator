package com.extia.fdaprocessor;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.GradientPaint;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.Point;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.IOException;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.SwingUtilities;

import org.apache.log4j.Logger;

import com.extia.fdaprocessor.io.ExcelIO;
import com.extia.fdaprocessor.io.FDAProcessUserSettingsIO;
import com.extia.fdaprocessor.io.FicheAductionIO;
import com.extia.fdaprocessor.ui.fdaprocessor.GUIFDAProcessor;

public class FDAProcessorLauncher {

	static Logger logger = Logger.getLogger(FDAProcessorLauncher.class);

	public void launch(String configFilePath) throws Exception{

		FicheAductionIO ficheIO = new FicheAductionIO();
		ficheIO.setExcelIO(new ExcelIO());

		final FDAProcessUserSettingsIO settingsIO = new FDAProcessUserSettingsIO();
		settingsIO.setConfigFilePath(configFilePath);

		final FDAProcessor fdaProcessor = new FDAProcessor();
		fdaProcessor.setFicheIO(ficheIO);
		fdaProcessor.setSettings(settingsIO.readScrappingSettings());

		GUIFDAProcessor guiFDAProcessor = new GUIFDAProcessor();
		guiFDAProcessor.setFdaProcessor(fdaProcessor);

		final JPanel contentPane = new JPanel(new GridBagLayout()){

			private static final long serialVersionUID = 6749083721571321159L;

			protected void paintComponent(Graphics g) {
				Color color1 = new Color(231, 248, 252);
				Color color2 = new Color(17, 156, 190);

				GradientPaint gradient = new GradientPaint(new Point(0, 0), color1, new Point(0, getHeight()), color2);

				((Graphics2D)g).setPaint(gradient);
				((Graphics2D)g).fill(getBounds());
			}
		};

		contentPane.add(guiFDAProcessor.getUI(), new GridBagConstraints(0, 0, 1, 1, 1, 1, GridBagConstraints.CENTER, GridBagConstraints.BOTH, new Insets(0, 0, 0, 0), 0, 0));

		
		SwingUtilities.invokeLater(new Runnable(){
			public void run() {
				JFrame frame = new JFrame();
				frame.setSize(new Dimension(800, 600));
				frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
				frame.getContentPane().add(contentPane);
				frame.setVisible(true);
				frame.addWindowListener(new WindowAdapter() {

					public void windowClosing(WindowEvent e) {
						try {
							settingsIO.writeScrappingSettings(fdaProcessor.getSettings());
						} catch (IOException ex) {
							logger.error(ex);
						}
					}
				});		
			}
		});
		
	}



	public static void main(String[] args) {
		try {
			String configFilePath = null;
			if(args != null){
				if(args.length > 0){
					configFilePath = args[0];
				}
			}

			FDAProcessorLauncher launcher = new FDAProcessorLauncher();
			launcher.launch(configFilePath);
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		}
	}

}
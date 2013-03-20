package com.extia.fdaprocessor;

import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.extia.fdaprocessor.data.Cable;
import com.extia.fdaprocessor.data.FDAProcessUserSettings;
import com.extia.fdaprocessor.data.FicheAduction;
import com.extia.fdaprocessor.data.Jaretiere;
import com.extia.fdaprocessor.io.ExcelIO;
import com.extia.fdaprocessor.io.FicheAductionIO;

public class FDAProcessor {

	static Logger logger = Logger.getLogger(FDAProcessor.class);

	private List<FDAProcessProgressListener> fDAProcessProgressListenerList;

	private FDAProcessUserSettings settings;
	
	private FicheAductionIO ficheIO;
	
	private int progressTotal;
	private int progressCurrent;
	
	public FDAProcessor() {
		fDAProcessProgressListenerList = new ArrayList<FDAProcessor.FDAProcessProgressListener>();
	}

	public int getProgressTotal() {
		return progressTotal;
	}

	public void setProgressTotal(int progressTotal) {
		this.progressTotal = progressTotal;
	}

	public int getProgressCurrent() {
		return progressCurrent;
	}

	public void setProgressCurrent(int progressCurrent) {
		this.progressCurrent = progressCurrent;
	}

	public FicheAductionIO getFicheIO() {
		return ficheIO;
	}

	public void setFicheIO(FicheAductionIO ficheIO) {
		this.ficheIO = ficheIO;
	}

	public boolean addFDAProcessProgressListener(FDAProcessProgressListener fDAProcessProgressListener){
		return fDAProcessProgressListenerList.add(fDAProcessProgressListener);
	}

	public boolean removeFDAProcessProgressListener(FDAProcessProgressListener fDAProcessProgressListener){
		return fDAProcessProgressListenerList.remove(fDAProcessProgressListener);
	}
	
	public void makeSyntheseFdas() throws InvalidFormatException, IOException{

		if(getSrcDir() != null && getDestDir() != null){

			List<FicheAduction> ficheList = readFicheList();

			if(ficheList != null){

				ExcelIO excelIO = getFicheIO().getExcelIO();
				
				HSSFWorkbook workbook = new HSSFWorkbook();
				
				
				List<Jaretiere> jaretiereBPIList = new ArrayList<Jaretiere>();
				for (FicheAduction ficheAduction : ficheList) {
					jaretiereBPIList.addAll(ficheAduction.getJaretiereBPIList());
				}

				List<Jaretiere> jaretiereNROList = new ArrayList<Jaretiere>();
				for (FicheAduction ficheAduction : ficheList) {
					jaretiereNROList.addAll(ficheAduction.getJaretiereNROList());
				}
				
				List<Cable> cableList = new ArrayList<Cable>();
				for (FicheAduction ficheAduction : ficheList) {
					cableList.addAll(ficheAduction.getCableList());
				}
				
				int progressTotal = cableList.size() + jaretiereBPIList.size() + jaretiereNROList.size();
				
				setProgressTotal(progressTotal);
				updateProgress(0);
				
				writeJaretiereList(workbook.createSheet("jaretières BPI"), jaretiereBPIList, excelIO);
				writeJaretiereList(workbook.createSheet("jaretières NRO"), jaretiereNROList, excelIO);
				writeCableList(workbook.createSheet("cables"), cableList, excelIO);

				FileOutputStream fileOutputStream = new FileOutputStream(new File(getDestDir(), "synthese_FDAs.xls"));
				workbook.write(fileOutputStream);
				fileOutputStream.close();
				
				finishProgress();
			}
		}
	}
	
	private void writeCableList(HSSFSheet sheet, List<Cable> cableList, ExcelIO excelIO) {
		if(sheet != null && cableList != null){
			int lineIndex = 0;
			
			excelIO.setCellValue("Identifiant du site", lineIndex, 0, sheet, true);
			excelIO.setCellValue("Equipement", lineIndex, 1, sheet, true);
			excelIO.setCellValue("Slot", lineIndex, 2, sheet, true);
			excelIO.setCellValue("Port", lineIndex, 3, sheet, true);
			excelIO.setCellValue("Position NRO", lineIndex, 4, sheet, true);
			excelIO.setCellValue("Cable de distribution", lineIndex, 5, sheet, true);
			excelIO.setCellValue("Couleur tube", lineIndex, 6, sheet, true);
			excelIO.setCellValue("Fibre", lineIndex, 7, sheet, true);
			excelIO.setCellValue("Couleur fibre", lineIndex, 8, sheet, true);
			excelIO.setCellValue("Splitter", lineIndex, 9, sheet, true);
			excelIO.setCellValue("Tray", lineIndex, 10, sheet, true);
			excelIO.setCellValue("Couleur fibre 2", lineIndex, 11, sheet, true);
			excelIO.setCellValue("Fibre 2", lineIndex, 12, sheet, true);
			excelIO.setCellValue("Couleur tube 2", lineIndex, 13, sheet, true);
			excelIO.setCellValue("Cable de raccordement", lineIndex, 14, sheet, true);
			lineIndex++;

			for (Cable cable : cableList) {
				excelIO.setCellValue(cable.getIdentifiantSite(), lineIndex, 0, sheet, true);
				excelIO.setCellValue(cable.getEquipement(), lineIndex, 1, sheet, true);
				excelIO.setCellValue(cable.getSlot(), lineIndex, 2, sheet, true);
				excelIO.setCellValue(cable.getPortFormatted(), lineIndex, 3, sheet, true);
				excelIO.setCellValue(cable.getPositionAuNRO(), lineIndex, 4, sheet, true);
				excelIO.setCellValue(cable.getCableDeDistrib(), lineIndex, 5, sheet, true);
				excelIO.setCellValue(cable.getCouleurTube(), lineIndex, 6, sheet, true);
				excelIO.setCellValue(cable.getFibreFormatted(), lineIndex, 7, sheet, true);
				excelIO.setCellValue(cable.getCouleurFibre(), lineIndex, 8, sheet, true);
				excelIO.setCellValue(cable.getSplitter(), lineIndex, 9, sheet, true);
				excelIO.setCellValue(cable.getTray(), lineIndex, 10, sheet, true);
				excelIO.setCellValue(cable.getCouleurFibre2(), lineIndex, 11, sheet, true);
				excelIO.setCellValue(cable.getFibre2Formatted(), lineIndex, 12, sheet, true);
				excelIO.setCellValue(cable.getCouleurTube2(), lineIndex, 13, sheet, true);
				excelIO.setCellValue(cable.getCableRaccordement(), lineIndex, 14, sheet, true);
				incrProgress();
				lineIndex++;
			}
			for(int i = 0; i <= 14; ++i){
				sheet.autoSizeColumn(i);
			}		
		}
	}

	private void incrProgress() {
		updateProgress(getProgressCurrent() + 1);
	}

	private void updateProgress(int currentProgress) {
		setProgressCurrent(currentProgress);
		fireProgressUpdated(Math.round(((float)(getProgressCurrent()) / (float)getProgressTotal()) * 100));		
	}
	
	private void finishProgress() {
		updateProgress(getProgressTotal());
	}

	private void writeJaretiereList(HSSFSheet sheet, List<Jaretiere> jaretiereList, ExcelIO excelIO) {
		if(sheet != null && jaretiereList != null){
			int lineIndex = 0;

			excelIO.setCellValue("Identifiant du site", lineIndex, 0, sheet, true);
			excelIO.setCellValue("Tenant", lineIndex, 1, sheet, true);
			excelIO.setCellValue("Aboutissant", lineIndex, 2, sheet, true);
			excelIO.setCellValue("Etat", lineIndex, 3, sheet, true);
			excelIO.setCellValue("Ref", lineIndex, 4, sheet, true);
			excelIO.setCellValue("Commentaires", lineIndex, 5, sheet, true);
			lineIndex++;

			for (Jaretiere jaretiere : jaretiereList) {
				excelIO.setCellValue(jaretiere.getIdentifiantSite(), lineIndex, 0, sheet, true);
				excelIO.setCellValue(jaretiere.getTenant(), lineIndex, 1, sheet, true);
				excelIO.setCellValue(jaretiere.getAboutissant(), lineIndex, 2, sheet, true);
				excelIO.setCellValue(jaretiere.getEtat(), lineIndex, 3, sheet, true);
				excelIO.setCellValue(jaretiere.getRef(), lineIndex, 4, sheet, true);
				excelIO.setCellValue(jaretiere.getCommentaires(), lineIndex, 5, sheet, true);
				
				incrProgress();
				lineIndex++;
			}
			for(int i = 0; i <= 5; ++i){
				sheet.autoSizeColumn(i);
			}
		}
	}

	public void processFdas() throws InvalidFormatException, IOException{
		if(getSrcDir() != null && getDestDir() != null){
			FileFilter xLSXFileFilter = new FileFilter(){
				public boolean accept(File file) {
					return file != null && file.getName().endsWith("xlsx") || file.getName().endsWith("xls");
				}
			};

			FicheAductionIO ficheDAductionIO = getFicheIO();
			//			ficheDAductionIO.displayWorkbook(sheetTemplate);

			int fileIndex = 0;
			
			
			File srcDir = getSrcDir();

			File[] srcFileList = srcDir.listFiles(xLSXFileFilter);

			setProgressTotal(srcFileList.length);
			updateProgress(0);

			for (File srcFile : srcFileList) {

				FicheAduction fiche = ficheDAductionIO.readFiche(srcFile);

				if(fiche != null){
					//TODO build imageList
					List<File> imageList = null;
					File imgFolder = new File(srcFile.getParentFile(), fiche.getIdentifiantSite());
					if(imgFolder.exists() && imgFolder.isDirectory()){
						File[] imageFileList = imgFolder.listFiles(new FilenameFilter() {
							public boolean accept(File file, String fileName) {
								boolean result = false;
								if(fileName != null){
									String lowerCaseFileName = fileName.toLowerCase();
									result = lowerCaseFileName != null && lowerCaseFileName.endsWith(".png") || lowerCaseFileName.endsWith(".jpg") || lowerCaseFileName.endsWith(".gif");
								}
								return result;
							}
						});
						imageList = imageFileList != null ? Arrays.asList(imageFileList) : null;
					}else{
						logger.warn("Le répertoire " + imgFolder.getAbsolutePath() + " n'a pas été résolu : aucune photo ne sera importée pour " + fiche.getIdentifiantSite());
					}
					InputStream is =  getClass().getResourceAsStream("/template.xls");
					Workbook workbookTemplate = WorkbookFactory.create(is);
					Sheet sheetTemplate = workbookTemplate.getSheetAt(0);

					ficheDAductionIO.writeFiche(fiche, sheetTemplate, imageList);

					FileOutputStream fileOutputStream = null;
					try{
						fileOutputStream = new FileOutputStream(new File(getDestDir(), fiche.getIdentifiantSite() + ".xls"));
						workbookTemplate.write(fileOutputStream);
					}finally{
						if(fileOutputStream != null){
							fileOutputStream.close();
						}
					}
				}

				incrProgress();
			}
			
			finishProgress();
		}


		//		Workbook workbookImageDir = WorkbookFactory.create(new FileInputStream("C:/Users/Michael Cortes/Desktop/FDA/photoRepository.xlsx"));
		//		List<? extends PictureData> lst = ((XSSFWorkbook)workbookImageDir).getAllPictures();
		//		for (Iterator<? extends PictureData> it = lst.iterator(); it.hasNext(); ) {
		//			XSSFPictureData pict = (XSSFPictureData)it.next();
		//
		//			String ext = pict.suggestFileExtension();
		//			byte[] data = pict.getData();
		//			
		//			
		//			InputStream in = new ByteArrayInputStream(data);
		//			BufferedImage bImageFromConvert = ImageIO.read(in);
		//			bImageFromConvert.getHeight();
		//
		//
		//			System.out.println("\npackage : " + pict.getPackagePart());
		//			System.out.println("relationship : " + pict.getPackageRelationship()); // null
		//			System.out.println("relations : " + pict.getRelations());//vide
		//			System.out.println("parent : " + pict.getParent());//null
		//			System.out.println(pict.getPackagePart().getRelationships());
		//			
		//			for (POIXMLDocumentPart xmlPart : pict.getRelations()) {
		//				System.out.println("  " + xmlPart.getClass());
		//			}
		//			System.out.println(pict.getClass() + "  ext ->" + ext);
		//			System.out.println(bImageFromConvert.getWidth() + ", " + bImageFromConvert.getHeight());
		//		}
	}


	private List<FicheAduction> readFicheList() throws InvalidFormatException, IOException {
		List<FicheAduction> result = null;

		if(getSrcDir() != null && getDestDir() != null){
			FileFilter xLSXFileFilter = new FileFilter(){
				public boolean accept(File file) {
					return file != null && file.getName().endsWith("xlsx") || file.getName().endsWith("xls");
				}
			};

			FicheAductionIO ficheDAductionIO = getFicheIO();
			//			ficheDAductionIO.displayWorkbook(sheetTemplate);

			fireProgressUpdated(0);

			File srcDir = getSrcDir();

			File[] srcFileList = srcDir.listFiles(xLSXFileFilter);
			
			result = new ArrayList<FicheAduction>();
			
			for (File srcFile : srcFileList) {

				FicheAduction fiche = ficheDAductionIO.readFiche(srcFile);

				if(fiche != null){
					result.add(fiche);
				}
			}
		}
		return result;
	}

	private void fireProgressUpdated(int progress){
		for (FDAProcessProgressListener fDAProcessProgressListener : fDAProcessProgressListenerList) {
			fDAProcessProgressListener.progressUpdated(progress);
		}
	}

	public interface FDAProcessProgressListener{
		void progressUpdated(int progress);
	}

	public File getSrcDir() {
		return getSettings().getSrcDir() != null ? new File(getSettings().getSrcDir()) : null;
	}

	public void setSrcDir(File srcDir) {
		getSettings().setSrcDir(srcDir != null ? srcDir.getAbsolutePath() : null);
	}

	public File getDestDir() {
		return getSettings().getDestDir() != null ? new File(getSettings().getDestDir()) : null;
	}

	public void setDestDir(File destDir) {
		getSettings().setDestDir(destDir != null ? destDir.getAbsolutePath() : null);
	}

	public void setSettings(FDAProcessUserSettings settings) {
		this.settings = settings;
	}

	public FDAProcessUserSettings getSettings() {
		return settings;
	}

}

package com.extia.fdaprocessor.excel;

import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.extia.fdaprocessor.data.FDAProcessUserSettings;
import com.extia.fdaprocessor.data.FicheDAduction;
import com.extia.fdaprocessor.io.FicheDAductionIO;

public class FDAProcessor {
	
	private List<FDAProcessProgressListener> fDAProcessProgressListenerList;
	private FDAProcessUserSettings settings;
	
	public FDAProcessor() {
		fDAProcessProgressListenerList = new ArrayList<FDAProcessor.FDAProcessProgressListener>();
	}
	
	public boolean addFDAProcessProgressListener(FDAProcessProgressListener fDAProcessProgressListener){
		return fDAProcessProgressListenerList.add(fDAProcessProgressListener);
	}

	public boolean removeFDAProcessProgressListener(FDAProcessProgressListener fDAProcessProgressListener){
		return fDAProcessProgressListenerList.remove(fDAProcessProgressListener);
	}

	public void processFdas() throws InvalidFormatException, IOException{
		if(getSrcDir() != null && getDestDir() != null){
			FileFilter xLSXFileFilter = new FileFilter(){
				public boolean accept(File file) {
					return file != null && file.getName().endsWith("xlsx");
				}
			};
			
			FicheDAductionIO ficheDAductionIO = new FicheDAductionIO();
			
//			ficheDAductionIO.displayWorkbook(sheetTemplate);
			
			int fileIndex = 0;
			fireProgressUpdated(0);
			
			File srcDir = getSrcDir();
			File destDir = getDestDir();
			
			File[] srcFileList = srcDir.listFiles(xLSXFileFilter);
			
			for (File srcFile : srcFileList) {
				
				FicheDAduction fiche = ficheDAductionIO.readFiche(srcFile);
				
				if(fiche != null){
					//TODO build imageList
					List<File> imageList = null;
//					File fileFolder = new File("C:/Users/Michael Cortes/Desktop/FDA");
//					File[] imageFileList = fileFolder.listFiles(new FilenameFilter() {
//						public boolean accept(File file, String fileName) {
//							boolean result = false;
//							if(fileName != null){
//								String lowerCaseFileName = fileName.toLowerCase();
//								result = lowerCaseFileName != null && lowerCaseFileName.endsWith(".png") || lowerCaseFileName.endsWith(".jpg") || lowerCaseFileName.endsWith(".gif");
//							}
//							return result;
//						}
//					});
					
					InputStream is =  getClass().getResourceAsStream("/template.xls");
					Workbook workbookTemplate = WorkbookFactory.create(is);
					Sheet sheetTemplate = workbookTemplate.getSheetAt(0);
					
					ficheDAductionIO.writeFiche(fiche, sheetTemplate, imageList);

					FileOutputStream fileOutputStream = new FileOutputStream(new File(destDir, fiche.getIdentifiantSite() + ".xls"));
					workbookTemplate.write(fileOutputStream);
					fileOutputStream.close();
				}
				
				int progress = Math.round(((float)(fileIndex ++) / (float)srcFileList.length) * 100);
				
				fireProgressUpdated(progress);
			}
			
			fireProgressUpdated(100);
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

package com.extia.fdaprocessor;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.extia.fdaprocessor.data.FicheDAduction;
import com.extia.fdaprocessor.io.FicheDAductionIO;

public class FDAProcessor {

	public void processFdas(File srcDir, File destDir) throws InvalidFormatException, IOException{
		if(srcDir != null && destDir != null){
			FileFilter xLSXFileFilter = new FileFilter(){

				public boolean accept(File file) {
					return file != null && file.getName().endsWith("xlsx");
				}
				
			};
			
			FicheDAductionIO ficheDAductionIO = new FicheDAductionIO();
			
			

//			ficheDAductionIO.displayWorkbook(sheetTemplate);

			for (File srcFile : srcDir.listFiles(xLSXFileFilter)) {
				InputStream is =  this.getClass().getResourceAsStream("/template.xls");
				Workbook workbookTemplate = WorkbookFactory.create(is);
				Sheet sheetTemplate = workbookTemplate.getSheetAt(0);
				
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
					
					ficheDAductionIO.writeFiche(fiche, sheetTemplate, imageList);

					FileOutputStream fileOutputStream = new FileOutputStream(new File(destDir, fiche.getIdentifiantSite() + ".xls"));
					workbookTemplate.write(fileOutputStream);
					fileOutputStream.close();
				}
			}
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
}

package com.extia.fdaprocessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.extia.fdaprocessor.data.FicheDAduction;
import com.extia.fdaprocessor.io.FicheDAductionIO;

public class FDAProcessorLauncher {

	public static void main(String[] args) throws Exception, FileNotFoundException, IOException {
		FicheDAductionIO ficheDAductionIO = new FicheDAductionIO();
		
		File fdaFile = new File("C:/Users/Michael Cortes/Desktop/FDA/FI-06088-002Y.xlsx");

		FicheDAduction fiche = ficheDAductionIO.readFiche(fdaFile);

		System.out.println(fiche);

		
		Workbook workbookTemplate = WorkbookFactory.create(new FileInputStream("C:/Users/Michael Cortes/Desktop/FDA/template.xls"));
		Sheet sheetTemplate = workbookTemplate.getSheetAt(0);
		
		ficheDAductionIO.displayWorkbook(sheetTemplate);

		ficheDAductionIO.writeFiche(fiche, sheetTemplate);

		FileOutputStream newExcelFile = new FileOutputStream("C:/Users/Michael Cortes/Desktop/FDA/result.xls");

		if(workbookTemplate instanceof HSSFWorkbook){
			File fileFolder = new File("C:/Users/Michael Cortes/Desktop/FDA");
			File[] imageFileList = fileFolder.listFiles(new FilenameFilter() {
				public boolean accept(File file, String fileName) {
					boolean result = false;
					if(fileName != null){
						String lowerCaseFileName = fileName.toLowerCase();
						result = lowerCaseFileName != null && lowerCaseFileName.endsWith(".png") || lowerCaseFileName.endsWith(".jpg") || lowerCaseFileName.endsWith(".gif");
					}
					return result;
				}
			});

			int rowIndex = 0;
			for (File imgFile : imageFileList) {
				ficheDAductionIO.addImage(((HSSFWorkbook)workbookTemplate).getSheetAt(1), imgFile, rowIndex++);
			}

		}
		workbookTemplate.write(newExcelFile);
		newExcelFile.close();
	}

}
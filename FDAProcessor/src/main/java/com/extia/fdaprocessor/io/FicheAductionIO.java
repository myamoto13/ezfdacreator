package com.extia.fdaprocessor.io;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import com.extia.fdaprocessor.data.Cable;
import com.extia.fdaprocessor.data.FicheAduction;
import com.extia.fdaprocessor.data.Jaretiere;

public class FicheAductionIO {

	static Logger logger = Logger.getLogger(FicheAductionIO.class);
	
	private ExcelIO excelIO;
	
	public ExcelIO getExcelIO() {
		return excelIO;
	}

	public void setExcelIO(ExcelIO excelIO) {
		this.excelIO = excelIO;
	}

	public FicheAduction readFiche(File fdaFile) throws InvalidFormatException, IOException {
		FicheAduction result = null;
		FileInputStream fi = null;
		Workbook workbook = null;
		try{
			fi = new FileInputStream(fdaFile);
			workbook = WorkbookFactory.create(fdaFile);

		}finally{
			if(fi != null){
				fi.close();
			}
		}
		if(workbook != null) {
			
			Sheet sheet = workbook.getSheetAt(0);

			if (sheet != null) {
				
				String fileName = FilenameUtils.removeExtension(fdaFile.getName());
				
				if(sheet.getRow(8).getLastCellNum() > 20){
					result = readProcessedFDASheet(sheet, fileName);
				}else{
					result = readRawFDASheet(sheet, fileName);
				}

			}
		}
		return result;
	}

	private FicheAduction readProcessedFDASheet(Sheet sheet, String fileName) {
		FicheAduction result = null;
		if(sheet != null){
			
			int indexRowJaretNRO = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if ("JARRETIERES A POSER au NRO".equals(getCellValue(indexRow, 0, sheet))) {
					indexRowJaretNRO = indexRow;
					break;
				}
			}

			int indexRowJaretBPI = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if ("JARRETIERES A POSER au BPI".equals(getCellValue(indexRow, 0, sheet))) {
					indexRowJaretBPI = indexRow;
					break;
				}
			}
			
			result = new FicheAduction();

			result.setIdentifiantSite(getCellValue(8, 21, sheet));
			if(result.getIdentifiantSite() == null || result.getIdentifiantSite().equals("")){
				result.setIdentifiantSite(fileName);
			}
			
			result.setDescription(getCellValue(0, 0, sheet));
			result.setDescDerivation(getCellValue(8, 11, sheet));

			for (int indexRow = 10; indexRow < indexRowJaretBPI; indexRow++) {
				
				String equipement = getCellValue(indexRow, 1, sheet);
				String slot = getCellValue(indexRow, 2, sheet);
				String portStr = getCellValue(indexRow, 3, sheet);
				String positionAuNRO = getCellValue(indexRow, 5, sheet);
				String cableDeDistrib = getCellValue(indexRow, 11, sheet);
				String couleurTube = getCellValue(indexRow, 12, sheet);
				String fibreStr = getCellValue(indexRow, 13, sheet);
				String couleurFibre = getCellValue(indexRow, 14, sheet);
				String splitter = getCellValue(indexRow, 15, sheet);
				String tray = getCellValue(indexRow, 16, sheet);
				String couleurFibre2 = getCellValue(indexRow, 17, sheet);
				String fibre2Str = getCellValue(indexRow, 18, sheet);
				String couleurTube2 = getCellValue(indexRow, 19, sheet);
				String cableRaccordement = getCellValue(indexRow, 20, sheet);

				Cable cable = createCable(equipement, slot, portStr,
						positionAuNRO, cableDeDistrib, couleurTube, fibreStr,
						couleurFibre, splitter, tray, couleurFibre2,
						fibre2Str, couleurTube2, cableRaccordement, result);
				if(cable != null){
					result.addCableList(cable);
				}
			}

			if (indexRowJaretBPI >= 0 && indexRowJaretBPI + 2 <= sheet.getLastRowNum()) {
				int maxRowJaretBPI = indexRowJaretNRO >= 0 ? indexRowJaretNRO - 1
						: sheet.getLastRowNum();
				for (int indexRow = indexRowJaretBPI + 2; indexRow < maxRowJaretBPI; indexRow++) {

					String tenant = getCellValue(indexRow, 0, sheet);
					String aboutissant = getCellValue(indexRow, 6, sheet);
					String ref = getCellValue(indexRow, 14, sheet);
					String etat = getCellValue(indexRow, 16, sheet);
					String commentaires = getCellValue(indexRow, 17, sheet);

					if (isValid(new String[] { tenant, aboutissant, ref, etat, commentaires })) {
						result.addJaretiereBPIList(createJaretiere(tenant, aboutissant, ref, etat, commentaires, result));
					}
				}

			}

			if (indexRowJaretNRO >= 0 && indexRowJaretNRO + 2 <= sheet.getLastRowNum()) {
				for (int indexRow = indexRowJaretNRO + 2; indexRow < sheet
						.getLastRowNum(); indexRow++) {

					String tenant = getCellValue(indexRow, 0, sheet);
					String aboutissant = getCellValue(indexRow, 6, sheet);
					String ref = getCellValue(indexRow, 14, sheet);
					String etat = getCellValue(indexRow, 16, sheet);
					String commentaires = getCellValue(indexRow, 17, sheet);

					if (isValid(new String[] { tenant, aboutissant, ref, etat, commentaires })) {
						result.addJaretiereNROList(createJaretiere(tenant, aboutissant, ref, etat, commentaires, result));
					}
				}

			}
		}
		return result;
	}
	
	private Cable createCable(String equipement, String slot, String portStr, String positionAuNRO, String cableDeDistrib, String couleurTube, String fibreStr, String couleurFibre, String splitter, String tray, String couleurFibre2, String fibre2Str, String couleurTube2, String cableRaccordement, FicheAduction fiche){
		Cable result = null;
		if (fiche != null && isValid(new String[] { equipement, slot, portStr,
				positionAuNRO, cableDeDistrib, couleurTube, fibreStr,
				couleurFibre, splitter, tray, couleurFibre2,
				fibre2Str, couleurTube2, cableRaccordement })) {
			Integer port = null;
			try{
				port = parseInt(portStr);
			}catch(Exception ex){
				logger.error("Format de port erroné pour " + fiche.getIdentifiantSite() + ".\n" + ex.getMessage());
			}

			Integer fibre = null;
			try{
				fibre = parseInt(fibreStr);
			}catch(Exception ex){
				logger.error("Format de fibre erroné pour " + fiche.getIdentifiantSite() + ".\n" + ex.getMessage());
			}

			Integer fibre2 = null;
			try{
				fibre2 = parseInt(fibre2Str);
			}catch(Exception ex){
				logger.error("Format de fibre2 erroné pour " + fiche.getIdentifiantSite() + ".\n" + ex.getMessage());
			}
			result = new Cable();
			result.setEquipement(equipement);
			result.setSlot(slot);
			result.setPort(port);
			result.setPositionAuNRO(positionAuNRO);
			result.setCableDeDistrib(cableDeDistrib);
			result.setCouleurTube(couleurTube);
			result.setFibre(fibre);
			result.setCouleurFibre(couleurFibre);
			result.setSplitter(splitter);
			result.setTray(tray);
			result.setCouleurFibre2(couleurFibre2);
			result.setFibre2(fibre2);
			result.setCouleurTube2(couleurTube2);
			result.setCableRaccordement(cableRaccordement);
			result.setFiche(fiche);
			
		}
		return result;
	}

	private FicheAduction readRawFDASheet(Sheet sheet, String fileName) {
		FicheAduction result = null;

		if(sheet != null){
			result = new FicheAduction();

			result.setIdentifiantSite(fileName);
			result.setDescription(getCellValue(0, 0, sheet));
			result.setDescDerivation(getCellValue(8, 5, sheet));

			for (int indexRow = 10; indexRow < sheet.getLastRowNum(); indexRow++) {

				String equipement = getCellValue(indexRow, 1, sheet);
				String slot = getCellValue(indexRow, 2, sheet);
				String portStr = getCellValue(indexRow, 3, sheet);
				String positionAuNRO = getCellValue(indexRow, 4, sheet);
				String cableDeDistrib = getCellValue(indexRow, 5, sheet);
				String couleurTube = getCellValue(indexRow, 6, sheet);
				String fibreStr = getCellValue(indexRow, 7, sheet);
				String couleurFibre = getCellValue(indexRow, 8, sheet);
				String splitter = getCellValue(indexRow, 9, sheet);
				String tray = getCellValue(indexRow, 10, sheet);
				String couleurFibre2 = getCellValue(indexRow, 11, sheet);
				String fibre2Str = getCellValue(indexRow, 12, sheet);
				String couleurTube2 = getCellValue(indexRow, 13, sheet);
				String cableRaccordement = getCellValue(indexRow, 14, sheet);

				if (!isValid(new String[] { equipement, slot, portStr,
						positionAuNRO, cableDeDistrib, couleurTube, fibreStr,
						couleurFibre, splitter, tray, couleurFibre2,
						fibre2Str, couleurTube2, cableRaccordement })) {
					break;
				}
				Integer port = null;
				try{
					port = parseInt(portStr);
				}catch(Exception ex){
					logger.error("Format de port erroné pour " + result.getIdentifiantSite() + ".\n" + ex.getMessage());
				}

				Integer fibre = null;
				try{
					fibre = parseInt(fibreStr);
				}catch(Exception ex){
					logger.error("Format de fibre erroné pour " + result.getIdentifiantSite() + ".\n" + ex.getMessage());
				}

				Integer fibre2 = null;
				try{
					fibre2 = parseInt(fibre2Str);
				}catch(Exception ex){
					logger.error("Format de fibre2 erroné pour " + result.getIdentifiantSite() + ".\n" + ex.getMessage());
				}

				Cable cable = new Cable();
				cable.setEquipement(equipement);
				cable.setSlot(slot);
				cable.setPort(port);
				cable.setPositionAuNRO(positionAuNRO);
				cable.setCableDeDistrib(cableDeDistrib);
				cable.setCouleurTube(couleurTube);
				cable.setFibre(fibre);
				cable.setCouleurFibre(couleurFibre);
				cable.setSplitter(splitter);
				cable.setTray(tray);
				cable.setCouleurFibre2(couleurFibre2);
				cable.setFibre2(fibre2);
				cable.setCouleurTube2(couleurTube2);
				cable.setCableRaccordement(cableRaccordement);
				cable.setFiche(result);
				result.addCableList(cable);
			}

			int indexRowJaretNRO = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if ("JARRETIERES A POSER au NRO".equals(getCellValue(indexRow, 0, sheet))) {
					indexRowJaretNRO = indexRow;
					break;
				}
			}

			int indexRowJaretBPI = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if ("JARRETIERES A POSER au BPI".equals(getCellValue(indexRow, 0, sheet))) {
					indexRowJaretBPI = indexRow;
					break;
				}
			}

			if (indexRowJaretBPI >= 0 && indexRowJaretBPI + 2 <= sheet.getLastRowNum()) {
				int maxRowJaretBPI = indexRowJaretNRO >= 0 ? indexRowJaretNRO - 1
						: sheet.getLastRowNum();
				for (int indexRow = indexRowJaretBPI + 2; indexRow < maxRowJaretBPI; indexRow++) {

					String tenant = getCellValue(indexRow, 0, sheet);
					String aboutissant = getCellValue(indexRow, 5, sheet);
					String ref = getCellValue(indexRow, 9, sheet);
					String etat = getCellValue(indexRow, 11, sheet);
					String commentaires = getCellValue(indexRow, 12, sheet);

					if (!isValid(new String[] { tenant, aboutissant, ref, etat, commentaires })) {
						break;
					}

					result.addJaretiereBPIList(createJaretiere(tenant, aboutissant, ref, etat, commentaires, result));

				}

			}

			if (indexRowJaretNRO >= 0 && indexRowJaretNRO + 2 <= sheet.getLastRowNum()) {
				for (int indexRow = indexRowJaretNRO + 2; indexRow < sheet.getLastRowNum(); indexRow++) {

					String tenant = getCellValue(indexRow, 0, sheet);
					String aboutissant = getCellValue(indexRow, 5, sheet);
					String ref = getCellValue(indexRow, 9, sheet);
					String etat = getCellValue(indexRow, 11, sheet);
					String commentaires = getCellValue(indexRow, 12, sheet);

					if (!isValid(new String[] { tenant, aboutissant, ref,
							etat, commentaires })) {
						break;
					}

					result.addJaretiereNROList(createJaretiere(tenant, aboutissant, ref, etat, commentaires, result));
				}

			}
		}
		return result;
	}

	private boolean isValid(String[] excelFieldList) {
		return getExcelIO().isValid(excelFieldList);
	}

	private String getCellValue(int rowIndex, int colIndex, Sheet sheet) {
		return getExcelIO().getCellValue(rowIndex, colIndex, sheet);
	}
	
	private void setCellValue(String value, int rowIndex, int colIndex, Sheet sheet) {
		getExcelIO().setCellValue(value, rowIndex, colIndex, sheet);
	}

	private Integer parseInt(String portStr) {
		Integer result = null;

		Double portDouble = Double.parseDouble(portStr);
		result = portDouble != null ? portDouble.intValue() : null;

		return result;
	}

	public void writeFiche(FicheAduction fiche, Sheet sheet, List<File> imageFileList) throws IOException {
		if (fiche != null && sheet != null) {

			setCellValue(fiche.getDescription(), 0, 0, sheet);

			setCellValue(fiche.getDescDerivation(), 8, 11, sheet);
			setCellValue(fiche.getDescDerivation(), 4, 16, sheet);

			setCellValue(fiche.getIdentifiantSite(), 8, 21, sheet);

			int indexRowIntituleJaretNRO = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if ("JARRETIERES A POSER au NRO".equals(getCellValue(indexRow, 0, sheet))) {
					indexRowIntituleJaretNRO = indexRow;
					break;
				}
			}

			int indexRowIntituleJaretBPI = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if ("JARRETIERES A POSER au BPI".equals(getCellValue(indexRow, 0, sheet))) {
					indexRowIntituleJaretBPI = indexRow;
					break;
				}
			}

			int indexCable = 10;
			for (Cable cable : fiche.getCableList()) {
				setCellValue(cable.getEntreprise(), indexCable, 0, sheet);

				setCellValue(cable.getEquipement(), indexCable, 1, sheet);
				setCellValue(cable.getSlot(), indexCable, 2, sheet);
				setCellValue(cable.getPortFormatted(), indexCable, 3, sheet);

				setCellValue(cable.getPositionAuNRO(), indexCable, 5, sheet);
				setCellValue(cable.getCableDeDistrib(), indexCable, 11, sheet);
				setCellValue(cable.getCouleurTube(), indexCable, 12, sheet);
				setCellValue(cable.getFibreFormatted(), indexCable, 13, sheet);
				setCellValue(cable.getCouleurFibre(), indexCable, 14, sheet);
				setCellValue(cable.getSplitter(), indexCable, 15, sheet);
				setCellValue(cable.getTray(), indexCable, 16, sheet);
				setCellValue(cable.getCouleurFibre2(), indexCable, 17, sheet);
				setCellValue(cable.getFibre2Formatted(), indexCable, 18, sheet);
				setCellValue(cable.getCouleurTube2(), indexCable, 19, sheet);
				setCellValue(cable.getCableRaccordement(), indexCable, 20, sheet);
				indexCable++;
			}

			List<CellRangeAddress> cellRangeList = new ArrayList<CellRangeAddress>();
			
			int entrepriseColIndex = 0;
			int previousIndex = 10;
			
			String previousVal = getCellValue(previousIndex, entrepriseColIndex, sheet);
			
			int maxIndexCable = 13;

			for (indexCable = previousIndex + 1; indexCable <= maxIndexCable; indexCable ++) {
				
				String newVal = getCellValue(indexCable, entrepriseColIndex, sheet);
				
				if(indexCable == maxIndexCable){
					cellRangeList.add(new CellRangeAddress(previousIndex, indexCable, entrepriseColIndex, entrepriseColIndex));
				}else if(previousVal == null || !previousVal.equals(newVal)){
					if(previousIndex != indexCable){
						cellRangeList.add(new CellRangeAddress(previousIndex, indexCable - 1, entrepriseColIndex, entrepriseColIndex));
						previousVal = newVal;
						previousIndex = indexCable;
					}
				}

			}
			for (CellRangeAddress cellRange : cellRangeList) {
				sheet.addMergedRegion(cellRange);
			}
			
			if (indexRowIntituleJaretBPI >= 0 && indexRowIntituleJaretBPI + 2 <= sheet.getLastRowNum()) {
				int indexRowJaretBPI = indexRowIntituleJaretBPI + 2;
				for (Jaretiere jaretiere : fiche.getJaretiereBPIList()) {

					setCellValue(jaretiere.getTenant(), indexRowJaretBPI, 0, sheet);
					setCellValue(jaretiere.getAboutissant(), indexRowJaretBPI, 6, sheet);
					setCellValue(jaretiere.getRef(), indexRowJaretBPI, 14, sheet);
					setCellValue(jaretiere.getEtat(), indexRowJaretBPI, 16, sheet);
					setCellValue(jaretiere.getCommentaires(), indexRowJaretBPI, 17, sheet);

					indexRowJaretBPI++;
				}

			}

			if (indexRowIntituleJaretNRO >= 0 && indexRowIntituleJaretNRO + 2 <= sheet.getLastRowNum()) {
				int indexRowJaretNRO = indexRowIntituleJaretNRO + 2;
				for (Jaretiere jaretiere : fiche.getJaretiereNROList()) {

					setCellValue(jaretiere.getTenant(), indexRowJaretNRO, 0, sheet);
					setCellValue(jaretiere.getAboutissant(), indexRowJaretNRO, 6, sheet);
					setCellValue(jaretiere.getRef(), indexRowJaretNRO, 14, sheet);
					setCellValue(jaretiere.getEtat(), indexRowJaretNRO, 16, sheet);
					setCellValue(jaretiere.getCommentaires(), indexRowJaretNRO, 17, sheet);

					indexRowJaretNRO++;
				}
			}
			
			if (imageFileList != null) {
				Workbook workbook = sheet.getWorkbook();
				if (workbook instanceof HSSFWorkbook) {
					int rowIndex = 0;
					for (File imgFile : imageFileList) {
						addImage(((HSSFWorkbook) workbook).getSheetAt(1), imgFile, rowIndex++);
					}

				}
			}
		}
	}

	private Jaretiere createJaretiere(String tenant, String aboutissant, String ref, String etat, String commentaires, FicheAduction fiche) {
		Jaretiere result = new Jaretiere();
		result.setTenant(tenant);
		result.setAboutissant(aboutissant);
		result.setRef(ref);
		result.setEtat(etat);
		result.setCommentaires(commentaires);
		result.setFiche(fiche);
		return result;
	}

	public void addImage(HSSFSheet sheet, File imageFile, int rowIndex) throws IOException {
			if (sheet != null) {
				// add picture data to this workbook.
				InputStream is = new FileInputStream(imageFile);
	
				byte[] bytes = IOUtils.toByteArray(is);
				int pictureIdx = sheet.getWorkbook().addPicture(bytes,
						Workbook.PICTURE_TYPE_JPEG);
				is.close();
	
				CreationHelper helper = sheet.getWorkbook().getCreationHelper();
	
				// Create the drawing patriarch. This is the top level container for
				// all shapes.
				if (sheet.getDrawingPatriarch() == null) {
					sheet.createDrawingPatriarch();
				}
	
				Drawing drawing = sheet.getDrawingPatriarch();
	
				// add a picture shape
				ClientAnchor anchor = helper.createClientAnchor();
				// set top-left corner of the picture,
				// subsequent call of Picture#resize() will operate relative to it
				anchor.setCol1(0);
				anchor.setRow1(rowIndex);
	
				setCellValue("" + rowIndex, rowIndex, 0, sheet);
	
				is = new FileInputStream(imageFile);
				BufferedImage img = ImageIO.read(is);
				is.close();
				int height = img.getHeight();
	
				HSSFRow row = sheet.getRow(rowIndex);
				if (row == null) {
					row = sheet.createRow(rowIndex);
				}
				row.setHeightInPoints(1 + height * 0.75F);
	
				Picture pict = drawing.createPicture(anchor, pictureIdx);
	
				// auto-size picture relative to its top-left corner
				pict.resize();
				
				if(!imageFile.getName().endsWith("gif")){
					logger.warn("Le format de l'image " + imageFile.getAbsolutePath() + " n'étant pas 'GIF', la taille et le positionnement ne sont pas correctement gérées.");
				}
				
//				System.out.println(imageFile);
//				System.out.println("x : " + pict.getPreferredSize().getDx1() + "  "
//						+ pict.getPreferredSize().getDx2());
//				System.out.println("y : " + pict.getPreferredSize().getDy1() + "  "
//						+ pict.getPreferredSize().getDy2());
	
			}
		}

	//	private void displayWorkbook(Sheet sheet) {
	//		for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
	//			Row row = sheet.getRow(indexRow);
	//			if (row != null) {
	//				for (int indexCol = 0; indexCol < row.getLastCellNum(); indexCol++) {
	//					Cell cell = row.getCell(indexCol);
	//					if (cell != null) {
	//						String cellVal = getCellValue(cell);
	//
	//						if (cellVal != null && !"".equals(cellVal)) {
	//							System.out.println("(" + indexRow + ", " + indexCol
	//									+ ") => " + cellVal);
	//						}
	//					}
	//				}
	//			}
	//		}
	//	}
}

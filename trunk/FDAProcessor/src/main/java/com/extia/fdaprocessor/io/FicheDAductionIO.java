package com.extia.fdaprocessor.io;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.imageio.ImageIO;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import com.extia.fdaprocessor.data.Cable;
import com.extia.fdaprocessor.data.FicheDAduction;
import com.extia.fdaprocessor.data.Jaretiere;

public class FicheDAductionIO {

	public FicheDAduction readFiche(File fdaFile) throws InvalidFormatException, IOException{
		FicheDAduction result = null;

		Workbook workbook = WorkbookFactory.create(fdaFile);

		if(workbook != null){

			Sheet sheet = workbook.getSheetAt(0);

			if(sheet != null){

				result = new FicheDAduction();

				result.setIdentifiantSite(FilenameUtils.removeExtension(fdaFile.getName()));
				result.setDescription(getCellValue(0, 0, sheet));
				result.setDescDerivation(getCellValue(8, 5, sheet));

				for (int indexRow = 10; indexRow < sheet.getLastRowNum(); indexRow++) {

					String equipement = getCellValue(indexRow, 1, sheet);
					String slot = getCellValue(indexRow, 2, sheet);
					String port = getCellValue(indexRow, 3, sheet);
					String positionAuNRO = getCellValue(indexRow, 4, sheet);
					String cableDeDistrib = getCellValue(indexRow, 5, sheet);
					String couleurTube = getCellValue(indexRow, 6, sheet);
					String fibre = getCellValue(indexRow, 7, sheet);
					String couleurFibre = getCellValue(indexRow, 8, sheet);
					String splitter = getCellValue(indexRow, 9, sheet);
					String tray = getCellValue(indexRow, 10, sheet);
					String couleurFibre2 = getCellValue(indexRow, 11, sheet);
					String fibre2 = getCellValue(indexRow, 12, sheet);
					String couleurTube2 = getCellValue(indexRow, 13, sheet);
					String cableRaccordement = getCellValue(indexRow, 14, sheet);

					if(!isValid(new String[]{equipement, slot, port, positionAuNRO, cableDeDistrib,couleurTube,
							fibre, couleurFibre, splitter, tray, couleurFibre2, fibre2, couleurTube2, cableRaccordement})){
						break;
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

					result.addCableList(cable);
				}

				int indexRowJaretNRO = -1;
				for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
					if("JARRETIERES A POSER au NRO".equals(getCellValue(indexRow, 0, sheet))){
						indexRowJaretNRO = indexRow;
						break;
					}
				}

				int indexRowJaretBPI = -1;
				for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
					if("JARRETIERES A POSER au BPI".equals(getCellValue(indexRow, 0, sheet))){
						indexRowJaretBPI = indexRow;
						break;
					}
				}

				if(indexRowJaretBPI >= 0 && indexRowJaretBPI + 2 <= sheet.getLastRowNum()){
					int maxRowJaretBPI = indexRowJaretNRO >= 0 ? indexRowJaretNRO - 1 : sheet.getLastRowNum();
					for (int indexRow = indexRowJaretBPI + 2; indexRow < maxRowJaretBPI; indexRow++) {

						String tenant = getCellValue(indexRow, 0, sheet);
						String aboutissant = getCellValue(indexRow, 5, sheet);
						String ref = getCellValue(indexRow, 9, sheet);
						String etat= getCellValue(indexRow, 11, sheet);
						String commentaires= getCellValue(indexRow, 12, sheet);

						if(!isValid(new String[]{tenant, aboutissant, ref, etat, commentaires})){
							break;
						}

						result.addJaretiereBPIList(createJaretiere(tenant, aboutissant, ref, etat, commentaires));

					}

				}

				if(indexRowJaretNRO >= 0 && indexRowJaretNRO + 2 <= sheet.getLastRowNum()){
					for (int indexRow = indexRowJaretNRO + 2; indexRow < sheet.getLastRowNum(); indexRow++) {

						String tenant = getCellValue(indexRow, 0, sheet);
						String aboutissant = getCellValue(indexRow, 5, sheet);
						String ref = getCellValue(indexRow, 9, sheet);
						String etat= getCellValue(indexRow, 11, sheet);
						String commentaires= getCellValue(indexRow, 12, sheet);

						if(!isValid(new String[]{tenant, aboutissant, ref, etat, commentaires})){
							break;
						}

						result.addJaretiereNROList(createJaretiere(tenant, aboutissant, ref, etat, commentaires));
					}

				}

			}
		}
		return result;
	}

	public void writeFiche(FicheDAduction fiche, Sheet sheet) {
		if(fiche != null && sheet != null){

			setCellValue(fiche.getDescription(), 0, 0, sheet);

			setCellValue(fiche.getDescDerivation(), 8, 11, sheet);
			setCellValue(fiche.getDescDerivation(), 4, 16, sheet);

			setCellValue(fiche.getIdentifiantSite(), 8, 21, sheet);

			int indexRowIntituleJaretNRO = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if("JARRETIERES A POSER au NRO".equals(getCellValue(indexRow, 0, sheet))){
					indexRowIntituleJaretNRO = indexRow;
					break;
				}
			}

			int indexRowIntituleJaretBPI = -1;
			for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
				if("JARRETIERES A POSER au BPI".equals(getCellValue(indexRow, 0, sheet))){
					indexRowIntituleJaretBPI = indexRow;
					break;
				}
			}


			int indexCable = 10;
			for (Cable cable : fiche.getCableList()) {

				setCellValue(cable.getEquipement(), indexCable, 1, sheet);
				setCellValue(cable.getSlot(), indexCable, 2, sheet);
				setCellValue(cable.getPort(), indexCable, 3, sheet);

				setCellValue(cable.getPositionAuNRO(), indexCable, 5, sheet);
				setCellValue(cable.getCableDeDistrib(), indexCable, 11, sheet);
				setCellValue(cable.getCouleurTube(), indexCable, 12, sheet);
				setCellValue(cable.getFibre(), indexCable, 13, sheet);
				setCellValue(cable.getCouleurFibre(), indexCable, 14, sheet);
				setCellValue(cable.getSplitter(), indexCable, 15, sheet);
				setCellValue(cable.getTray(), indexCable, 16, sheet);
				setCellValue(cable.getCouleurFibre2(), indexCable, 17, sheet);
				setCellValue(cable.getFibre2(), indexCable, 18, sheet);
				setCellValue(cable.getCouleurTube2(), indexCable, 19, sheet);
				setCellValue(cable.getCableRaccordement(), indexCable, 20, sheet);
				indexCable++;
			}

			if(indexRowIntituleJaretBPI >= 0 && indexRowIntituleJaretBPI + 2 <= sheet.getLastRowNum()){
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

			if(indexRowIntituleJaretNRO >= 0 && indexRowIntituleJaretNRO + 2 <= sheet.getLastRowNum()){
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
		}
	}

	private Jaretiere createJaretiere(String tenant, String aboutissant, String ref, String etat, String commentaires) {
		Jaretiere result = new Jaretiere();
		result.setTenant(tenant);
		result.setAboutissant(aboutissant);
		result.setRef(ref);
		result.setEtat(etat);
		result.setCommentaires(commentaires);

		return result;
	}

	private boolean isValid(String... excelFieldList) {
		boolean result = false;
		if(excelFieldList != null){
			for (int i = 0; i < excelFieldList.length; i++) {
				if(excelFieldList[i] != null && !"".equals(excelFieldList[i])){
					result = true;
					break;
				}
			}
		}
		return result;
	}

	private void setCellValue(String value, int rowIndex, int colIndex, Sheet sheet) {
		Cell cell = getCell(rowIndex, colIndex, sheet);
		if(cell != null){
			cell.setCellValue(value);
			System.out.println(value + "  "  + cell.getCellStyle().getWrapText());
		}
	}

	private Cell getCell(int rowIndex, int colIndex, Sheet sheet) {
		Cell result = null;
		if(sheet != null){
			Row row = sheet.getRow(rowIndex);
			if(row != null){
				result = row.getCell(colIndex);
			}
		}
		return result;
	}

	private String getCellValue(int rowIndex, int colIndex, Sheet sheet) {
		return getCellValue(getCell(rowIndex, colIndex, sheet));
	}

	private String getCellValue(Cell cell) {
		String result = null;
		if(cell != null){
			//TODO : check for nullpointers
			FormulaEvaluator formulaEv = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();

			CellValue cValue = formulaEv.evaluate(cell);
			switch(cell.getCellType()){
			case Cell.CELL_TYPE_BOOLEAN :
				result = "" + cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC :
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					Date date = HSSFDateUtil.getJavaDate(cValue.getNumberValue());
					String dateFmt = cell.getCellStyle().getDataFormatString();
					result = new SimpleDateFormat("dd/mm/YYYY").format(date);
					//					result = date.toString() + "  (" + dateFmt + ") ==> " + (new SimpleDateFormat(dateFmt).format(date));
				}else{
					result = "" + cell.getNumericCellValue();
				}
				break;
			case Cell.CELL_TYPE_ERROR :
				result = "" + cell.getErrorCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA :
				result = "" + cell.getStringCellValue() + "(" + cell.getCellFormula() + ")";
				break;
			case Cell.CELL_TYPE_BLANK :
			case Cell.CELL_TYPE_STRING :
				result = cell.getStringCellValue();
				break;
			}
		}
		return result;
	}

	public void addImage(HSSFSheet sheet, File imageFile, int rowIndex) throws IOException {
		//create a new workbook
		if(sheet != null){
			//add picture data to this workbook.
			InputStream is = new FileInputStream(imageFile);

			byte[] bytes = IOUtils.toByteArray(is);
			int pictureIdx = sheet.getWorkbook().addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
			is.close();

			CreationHelper helper = sheet.getWorkbook().getCreationHelper();


			// Create the drawing patriarch. This is the top level container for all shapes. 
			if(sheet.getDrawingPatriarch() == null){
				sheet.createDrawingPatriarch();
			}

			Drawing drawing = sheet.getDrawingPatriarch();

			//add a picture shape
			ClientAnchor anchor = helper.createClientAnchor();
			//set top-left corner of the picture,
			//subsequent call of Picture#resize() will operate relative to it
			anchor.setCol1(0);
			anchor.setRow1(rowIndex);

			setCellValue("" + rowIndex, rowIndex, 0, sheet);


			is = new FileInputStream(imageFile);
			BufferedImage img = ImageIO.read(is);
			is.close();
			int height = img.getHeight();

			HSSFRow row = sheet.getRow(rowIndex);
			if(row == null){
				row = sheet.createRow(rowIndex);
			}
			row.setHeightInPoints(1 + height * 0.75F);


			Picture pict = drawing.createPicture(anchor, pictureIdx);

			//auto-size picture relative to its top-left corner
			pict.resize();

			System.out.println(imageFile);
			System.out.println("x : " + pict.getPreferredSize().getDx1() + "  " +  pict.getPreferredSize().getDx2());
			System.out.println("y : " + pict.getPreferredSize().getDy1() + "  " +  pict.getPreferredSize().getDy2());


		}
	}

	public void displayWorkbook(Sheet sheet) {
		for (int indexRow = 0; indexRow < sheet.getLastRowNum(); indexRow++) {
			Row row = sheet.getRow(indexRow);
			if(row != null){
				for (int indexCol = 0; indexCol < row.getLastCellNum(); indexCol++) {
					Cell cell = row.getCell(indexCol);
					if(cell != null){
						String cellVal = getCellValue(cell);

						if(cellVal != null && !"".equals(cellVal)){
							System.out.println("(" + indexRow + ", " + indexCol + ") => " + cellVal);
						}
					}
				}
			}
		}		
	}
}


package com.extia.fdaprocessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ddf.EscherBSERecord;
import org.apache.poi.ddf.EscherBlipRecord;
import org.apache.poi.ddf.EscherClientAnchorRecord;
import org.apache.poi.ddf.EscherRecord;
import org.apache.poi.hslf.usermodel.PictureData;
import org.apache.poi.hssf.record.AbstractEscherHolderRecord;
import org.apache.poi.hssf.record.EscherAggregate;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author win user
 */
public class WorkbookPictureReaderTest {

    public void getImageMatrices(String folderName) throws IOException, FileNotFoundException, InvalidFormatException {
        File fileFolder = new File(folderName);
        File[] excelFileList = fileFolder.listFiles(new ExcelFilenameFilter());
        for(File excelFile : excelFileList) {
            Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFile));
            if(workbook instanceof HSSFWorkbook) {
                this.processImages((HSSFWorkbook)workbook, excelFile);
            }
            else {
                this.processImages((XSSFWorkbook)workbook);
            }
        }
    }
    
    private void searchForPictures(List<EscherRecord> escherRecords, List<HSSFPictureData> pictures)
    {
        for(EscherRecord escherRecord : escherRecords) {

            if (escherRecord instanceof EscherBSERecord)
            {
                EscherBlipRecord blip = ((EscherBSERecord) escherRecord).getBlipRecord();
                if (blip != null)
                {
                    // TODO: Some kind of structure.
                    HSSFPictureData picture = new HSSFPictureData(blip);
					pictures.add(picture);
                }
                
                
            }

            // Recursive call.
            searchForPictures(escherRecord.getChildRecords(), pictures);
        }
    }

    private void processImages(HSSFWorkbook workbook, File excelFile) {
        EscherAggregate drawingAggregate = null;
        HSSFSheet sheet = null;
        List<EscherRecord> recordList = null;
        Iterator<EscherRecord> recordIter = null;
        
//        System.out.println("pictures");
//        for(HSSFPictureData pictureData : workbook.getAllPictures()){
//        	
//        }
        
        /*
         * TODO : impossible pour l'instant d'avoir les datas de l'image et de savoir de quelle worksheet l'image provient.
         * http://stackoverflow.com/questions/9232844/get-picture-position-in-apache-poi-from-excel-xls-hssf
         * http://apache-poi.1045710.n5.nabble.com/A-newbie-question-how-to-get-image-position-td2294815.html
         */
        // The drawing group record always exists at the top level, so we won't need to do this recursively.
//        List<HSSFPictureData> pictures = new ArrayList<HSSFPictureData>();
//        Iterator<Record> recordIter2 = workbook.workbook.getRecords().iterator();
//        while (recordIter2.hasNext())
//        {
//            Record r = recordIter2.next();
//            if (r instanceof AbstractEscherHolderRecord)
//            {
//                ((AbstractEscherHolderRecord) r).decode();
//                List<EscherRecord> escherRecords = ((AbstractEscherHolderRecord) r).getEscherRecords();
//                searchForPictures(escherRecords, pictures);
//            }
//        }

    
        
        for(int i = 0; i < workbook.getNumberOfSheets(); i++) {
            System.out.println("Processing " + excelFile.getName() + " sheet number: " + (i + 1));
            sheet = workbook.getSheetAt(i);
            drawingAggregate = sheet.getDrawingEscherAggregate();
            if(drawingAggregate != null) {
                recordList = drawingAggregate.getEscherRecords();
                recordIter = recordList.iterator();
                while(recordIter.hasNext()) {
                    this.iterateRecords(recordIter.next(), 1);
                }
            }
        }
    }

    private void iterateRecords(EscherRecord escherRecord, int level) {
        List<EscherRecord> recordList = null;
        Iterator<EscherRecord> recordIter = null;
        EscherRecord childRecord = null;
        recordList = escherRecord.getChildRecords();
        recordIter = recordList.iterator();
        while(recordIter.hasNext()) {
            childRecord = recordIter.next();
            if(childRecord instanceof EscherClientAnchorRecord) {
            	EscherClientAnchorRecord anchorRecord = (EscherClientAnchorRecord)childRecord;
            	
            	System.out.println("  " + anchorRecord.getRecordName());
            	System.out.println("  " + anchorRecord.getRecordId());
            	
            	System.out.println("  The top left hand corner of the image can be found " +
                        "in the cell at column number " +
                        anchorRecord.getCol1() +
                        " and row number " +
                        anchorRecord.getRow1() +
                        " at the offset position x " +
                        anchorRecord.getDx1() +
                        " and y " +
                        anchorRecord.getDy1() + 
                        " co-ordinates.");
                System.out.println("  The bottom right hand corner of the image can be found " +
                        "in the cell at column number " +
                        anchorRecord.getCol2() +
                        " and row number " +
                        anchorRecord.getRow2() +
                        " at the offset position x " +
                        anchorRecord.getDx2() +
                        " and y " +
                        anchorRecord.getDy2() +
                        " co-ordinates.");
            }
            if(childRecord.getChildRecords().size() > 0) {
                this.iterateRecords(childRecord, ++level);
            }
        }
    }


    private void processImages(XSSFWorkbook workbook) {
        System.out.println("No support yet for OOXML based workbooks. Investigating.");
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            new WorkbookPictureReaderTest().getImageMatrices("C:/Users/Michael Cortes/Desktop/FDA");
        }catch(Exception ex) {
            System.out.println("Caught an: " + ex.getClass().getName());
            System.out.println("Message: " + ex.getMessage());
            System.out.println("Stacktrace follows:.....");
            ex.printStackTrace(System.out);
        }
    }

    public class ExcelFilenameFilter implements FilenameFilter {
        public boolean accept(File file, String fileName) {
            return fileName.endsWith("photoRepository.xls");
        }
    }
} 
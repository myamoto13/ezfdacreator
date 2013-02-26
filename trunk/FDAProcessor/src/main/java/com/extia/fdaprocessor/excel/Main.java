package com.extia.fdaprocessor.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ddf.EscherClientAnchorRecord;
import org.apache.poi.ddf.EscherRecord;
import org.apache.poi.hssf.record.EscherAggregate;
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
public class Main {

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

    private void processImages(HSSFWorkbook workbook, File excelFile) {
        EscherAggregate drawingAggregate = null;
        HSSFSheet sheet = null;
        List<EscherRecord> recordList = null;
        Iterator<EscherRecord> recordIter = null;
        int numSheets = workbook.getNumberOfSheets();
        for(int i = 0; i < numSheets; i++) {
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
                this.printAnchorDetails((EscherClientAnchorRecord)childRecord);
            }
            if(childRecord.getChildRecords().size() > 0) {
                this.iterateRecords(childRecord, ++level);
            }
        }
    }

    private void printAnchorDetails(EscherClientAnchorRecord anchorRecord) {
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

    private void processImages(XSSFWorkbook workbook) {
        System.out.println("No support yet for OOXML based workbooks. Investigating.");
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            new Main().getImageMatrices("C:/Users/Michael Cortes/Desktop/FDA");
        }catch(Exception ex) {
            System.out.println("Caught an: " + ex.getClass().getName());
            System.out.println("Message: " + ex.getMessage());
            System.out.println("Stacktrace follows:.....");
            ex.printStackTrace(System.out);
        }
    }

    public class ExcelFilenameFilter implements FilenameFilter {
        public boolean accept(File file, String fileName) {
            return fileName.endsWith(".xls") || fileName.endsWith(".xlsx");
        }
    }
} 
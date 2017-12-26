package jp.co.japandisplay;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.math.BigInteger;
import java.math.BigDecimal;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ReadExcelExample {

    private static final String FILE_NAME = "./台帳マスタ.xlsx";

    public static void main(String[] args) { 
    	ReadExcelExample testFile = new ReadExcelExample();
    	testFile.readFromFile();
    	testFile.writeToFile();

    }
    
    public void readFromFile() {

        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            iterator.next();
            iterator.next();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.println(currentCell.getStringCellValue());
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        BigInteger intVal = new BigDecimal(currentCell.getNumericCellValue()).toBigInteger();
                        System.out.println(intVal);
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
    
    public void writeToFile() {
    	try {
	    	FileInputStream inp = new FileInputStream(new File(FILE_NAME));
	        Workbook wb = WorkbookFactory.create(inp);
	        Sheet sheet = wb.getSheetAt(0);
	        Row row = sheet.createRow(2);
	        Cell cell = row.getCell(6);
	        if (cell == null)
	            cell = row.createCell(3);
	        cell.setCellType(CellType.STRING);
	        cell.setCellValue("a test");
	
	        // Write the output to a file
	        FileOutputStream fileOut = new FileOutputStream(new File("test.xlsx"));
	        wb.write(fileOut);
	        fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
        	e.printStackTrace();
        }
    }    
}
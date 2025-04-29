package xlsx;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Base64;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class multipleExcelComparing {
    public static void main(String[] args) {
        try (FileInputStream fi1 = new FileInputStream("C:\\Users\\NICHEBIT\\Desktop\\Dummy1.xlsx");
             FileInputStream fi2 = new FileInputStream("C:\\Users\\NICHEBIT\\Desktop\\Dummy2.xlsx");
             XSSFWorkbook wb1 = new XSSFWorkbook(fi1);
             XSSFWorkbook wb2 = new XSSFWorkbook(fi2)) {

            Sheet sheet1 = wb1.getSheetAt(0);
            Sheet sheet2 = wb2.getSheetAt(0);

          
            int maxRowLen = Math.max(sheet1.getLastRowNum(), sheet2.getLastRowNum());

            for (int i1 = 0; i1 <= maxRowLen; i1++) {
                Row row1 = sheet1.getRow(i1);
                Row row2 = sheet2.getRow(i1);
                if(row1!=null && row2!=null) {
                	 for (int i = 0; i < row1.getLastCellNum(); i++) {
                         Cell cell1 = row1.getCell(i);
                         Cell cell2 = row2.getCell(i);

                         if (cell1 != null && cell2 != null && !cell1.toString().equals(cell2.toString())) {
                        		CellStyle cellStyle= cell2.getCellStyle();
                                cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                             cell2.setCellStyle(cellStyle);
                             System.out.println("Mismatch found at row " + row2.getRowNum() + ", column " + i);
                             System.out.println("Value in wb1: " + cell1.toString());
                             System.out.println("Value in wb2: " + cell2.toString());
                         }
                     }
                }
                else if(row2!=null && row1==null) {
                	for (int i = 0; i < row2.getLastCellNum(); i++) {
                        Cell cell2 = row2.getCell(i);
                        if ( cell2 != null) {
                        	CellStyle cellStyle= cell2.getCellStyle();
                            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell2.setCellStyle(cellStyle);
                        }
                    }	
                }
                else if(row2==null && row1!=null) {
                	for (int i = 0; i < row1.getLastCellNum(); i++) {
                        Cell cell2 = row1.getCell(i);
                        if ( cell2 != null) {
                        
                        	CellStyle cellStyle= cell2.getCellStyle();
                            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell2.setCellStyle(cellStyle);
                        }
                    }	
                }
            }
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                 FileOutputStream fileOut = new FileOutputStream("C:\\Users\\NICHEBIT\\Desktop\\Dummy1_updated1.xlsx")) {
                wb2.write(outputStream);
                fileOut.write(outputStream.toByteArray());
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

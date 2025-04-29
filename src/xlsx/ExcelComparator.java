package xlsx;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelComparator {
    public static void main(String[] args) throws IOException {
    	List<Integer> cellColWidth=new ArrayList<>();
        String filePath1 =  "C:\\Users\\NICHEBIT\\Downloads\\IVN_03-Sep-2024 114000.xlsx";
        String filePath2 = "C:\\Users\\NICHEBIT\\Downloads\\IVN_03-Sep-2024 115538.xlsx";
        String outputFilePath = "C:\\Users\\NICHEBIT\\Desktop\\Dummy1_updated.xlsx";

        
       
        
        FileInputStream file1 = new FileInputStream(filePath1);
        FileInputStream file2 = new FileInputStream(filePath2);

        Workbook workbook1 = new XSSFWorkbook(file1);
        Workbook workbook2 = new XSSFWorkbook(file2);
        Workbook outputWorkbook = new XSSFWorkbook();

        Sheet sheet1 = workbook1.getSheetAt(0);
        Sheet sheet2 = workbook2.getSheetAt(0);
        Sheet outputSheet = outputWorkbook.createSheet("Comparison");

        int maxRowNum = Math.max(sheet1.getLastRowNum(), sheet2.getLastRowNum());
        int maxCellNum=0;
        for (int i = 0; i <= maxRowNum; i++) {
            Row row1 = sheet1.getRow(i);
            Row row2 = sheet2.getRow(i);
            Row outputRow = outputSheet.createRow(i);
          
            if (i == 0) {
            	maxCellNum = Math.max(
                row1 != null ? row1.getLastCellNum() : 0, 
                row2 != null ? row2.getLastCellNum() : 0
            );
            }
           
            for (int j = 0; j < maxCellNum; j++) {
                Cell cell1 = row1 != null ? row1.getCell(j) : null;
                Cell cell2 = row2 != null ? row2.getCell(j) : null;
                Cell outputCell = outputRow.createCell(j);
				if (i == 0) {
						cellColWidth.add(Math.max(sheet2.getColumnWidth(j),sheet1.getColumnWidth(j)));
				}
                CellStyle style = outputWorkbook.createCellStyle();

                if (cell1 == null || cell2 == null) {
                    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    outputCell.setCellStyle(style);
                    outputCell.setCellValue(cell1 != null ? cell1.toString() : cell2 != null ?cell2.toString():"");
                } else if (!cell1.toString().equals(cell2.toString())) {
                    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    outputCell.setCellStyle(style);
                    outputCell.setCellValue(cell1.toString());
                } else {
                	CellStyle outputCellStyle = outputWorkbook.createCellStyle();
                	CellStyle sourceStyle=cell1!=null?cell1.getCellStyle():cell2.getCellStyle();
                    outputCell.setCellValue(cell1.toString());
                    outputCellStyle.cloneStyleFrom(sourceStyle);
                    outputCell.setCellStyle(outputCellStyle);
                }
            }
        }
        for(int i=0;i<cellColWidth.size();i++) {
        	outputSheet.setColumnWidth(i,cellColWidth.get(i));	
        }
        FileOutputStream outputStream = new FileOutputStream(outputFilePath);
        outputWorkbook.write(outputStream);
        outputStream.close();
        System.out.println(cellColWidth);
        workbook1.close();
        workbook2.close();
        outputWorkbook.close();
       
    }
}

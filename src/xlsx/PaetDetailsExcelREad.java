package xlsx;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PaetDetailsExcelREad {
	 public static boolean isInteger(String str) {
	        return str != null && str.matches("\\d+");
	    }
	public static void main(String[] args) throws Exception {

		List<PartDetails> partDetailsList = new ArrayList<>();
		partDetailsList.add(new PartDetails("BIN1", "PART1"));
		partDetailsList.add(new PartDetails("BIN2", "PART2"));
		partDetailsList.add(new PartDetails("BIN3", "PART3"));

		XSSFWorkbook workbook = null;
		String Filepath = "C:\\Users\\NICHEBIT\\Desktop\\UploadPartDetailsTemplate.xlsx";
		FileInputStream fis = new FileInputStream(Filepath);
		workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rows = sheet.iterator();
		int rowNo = 0;
		List<String> partDetailsAdded = new ArrayList<>();
		String outPut = "";
		while (rows.hasNext()) {
			boolean available = false;
			rowNo++;
			Row row = rows.next();
			Cell CurrentBinNo = row.getCell(2);
			
			
			 
			
			
			
			if (rowNo > 1) {
				int type=CurrentBinNo.getCellType();
				//double stringCellValue = CurrentBinNo.getNumericCellValue();
//				for (PartDetails element : partDetailsList) {
//					if (element.binNo.equals(stringCellValue)) {
//						available = true;
//					}
//				}
//				boolean contains = partDetailsAdded.contains(stringCellValue);
//				if (contains) {
//				} else {
//					//partDetailsAdded.add(stringCellValue);
//					if (available) {
//						outPut += stringCellValue + "         ";
//					}
//				}
			//	System.out.println(stringCellValue);
			}
		}
		System.out.println(outPut);
	}
}

class PartDetails {
	String binNo;
	String partNo;

	public PartDetails(String binNo, String partNo) {
		this.binNo = binNo;
		this.partNo = partNo;
	}
}

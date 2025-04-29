package xlsx;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.YearMonth;

import org.apache.poi.hssf.util.HSSFColor.YELLOW;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PMMExcelTemplate {

	public static void main(String[] args) throws Exception {
		
		int[] daysPerMonthInFinancialYear = PMMExcelTemplate.getDaysPerMonthInFinancialYear();
		String filePath = "C:\\Users\\NICHEBIT\\Desktop\\test\\Testing32.xlsx";
		
		
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		
//		sheet.protectSheet("password");
		
		
		XSSFFont font1 = workbook.createFont();
		font1.setBold(true);
		
		

		
		
		
		byte redLY = (byte) 255 ;
		byte greenLY = (byte) 193 ;
		byte blueLY = (byte) 7;
		
		byte redSC = (byte) 250 ;
		byte greenSC = (byte) 215 ;
		byte blueSC = (byte) 200;
		XSSFColor SkinColor = new XSSFColor(new byte[] { redSC, greenSC, blueSC });
		XSSFColor DarkYellow = new XSSFColor(new byte[] { redLY, greenLY, blueLY });
		
		XSSFCellStyle styleLY = (XSSFCellStyle) workbook.createCellStyle();
		styleLY.setBorderBottom(BorderStyle.THIN);
		styleLY.setBorderTop(BorderStyle.THIN);
		styleLY.setBorderRight(BorderStyle.THIN);
		styleLY.setBorderLeft(BorderStyle.THIN);
		styleLY.setWrapText(true);
		styleLY.setAlignment(HorizontalAlignment.CENTER);
		styleLY.setVerticalAlignment(VerticalAlignment.CENTER);
		styleLY.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		styleLY.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleLY.setFont(font1);
		styleLY.setLocked(true);
		
		
		XSSFCellStyle styleDY = (XSSFCellStyle) workbook.createCellStyle();
		styleDY.setBorderBottom(BorderStyle.THIN);
		styleDY.setBorderTop(BorderStyle.THIN);
		styleDY.setBorderRight(BorderStyle.THIN);
		styleDY.setBorderLeft(BorderStyle.THIN);
		styleDY.setWrapText(true);
		styleDY.setAlignment(HorizontalAlignment.CENTER);
		styleDY.setVerticalAlignment(VerticalAlignment.CENTER);
		styleDY.setFillForegroundColor(DarkYellow);
		styleDY.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleDY.setFont(font1);
		styleDY.setLocked(false);
		
		
		XSSFCellStyle stylef= (XSSFCellStyle) workbook.createCellStyle();
		stylef.setFont(font1);
		
		XSSFCellStyle styleb= (XSSFCellStyle) workbook.createCellStyle();
		styleb.setBorderBottom(BorderStyle.THIN);
		styleb.setBorderTop(BorderStyle.THIN);
		styleb.setBorderRight(BorderStyle.THIN);
		styleb.setBorderLeft(BorderStyle.THIN);
		
		XSSFCellStyle styleSC = (XSSFCellStyle) workbook.createCellStyle();
		styleSC.setBorderBottom(BorderStyle.THIN);
		styleSC.setBorderTop(BorderStyle.THIN);
		styleSC.setBorderRight(BorderStyle.THIN);
		styleSC.setBorderLeft(BorderStyle.THIN);
		styleSC.setWrapText(true);
		styleSC.setAlignment(HorizontalAlignment.CENTER);
		styleSC.setVerticalAlignment(VerticalAlignment.CENTER);
		styleSC.setFillForegroundColor(SkinColor);
		styleSC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleSC.setFont(font1);
		
		XSSFRow HeadderRow0= sheet.createRow(0);
		HeadderRow0.setHeightInPoints(30);
		XSSFRow HeadderRow1= sheet.createRow(1);
		HeadderRow1.setHeightInPoints(20);
		XSSFRow HeadderRow2= sheet.createRow(2);
		HeadderRow2.setHeightInPoints(20);
		XSSFRow HeadderRow3= sheet.createRow(3);
		HeadderRow3.setHeightInPoints(20);
		
		
		
		XSSFRow normalHeadder1= sheet.createRow(4);
		normalHeadder1.setHeightInPoints(20);
		XSSFRow normalHeadder2= sheet.createRow(5);
		normalHeadder2.setHeightInPoints(20);
		XSSFRow normalHeadder3= sheet.createRow(6);
		normalHeadder3.setHeightInPoints(20);
		XSSFRow normalHeadder4= sheet.createRow(7);
		normalHeadder4.setHeightInPoints(20);
		XSSFRow normalHeadder5= sheet.createRow(8);
		normalHeadder5.setHeightInPoints(20);
		XSSFRow normalHeadder6= sheet.createRow(9);
		normalHeadder6.setHeightInPoints(20);
		XSSFRow normalHeadder7= sheet.createRow(10);
		normalHeadder7.setHeightInPoints(20);
		XSSFRow normalHeadder8= sheet.createRow(11);
		normalHeadder8.setHeightInPoints(20);
		
		
		
		
		sheet.addMergedRegion(new CellRangeAddress(0, 3, 0, 1));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(0, 3, 0, 1), "");
		//addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\LogoForExcel.png", 0, 3, 0, 2, "");
		sheet.addMergedRegion(new CellRangeAddress(0, 3, 2, 12));
		HeadderRow0.createCell(2).setCellValue("Test lab Maintenance Equipment - Preventive Maintenance Yearly Plan ");
		
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(0, 3, 2, 12), "headding");
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 13, 18));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(0, 0, 13, 18), "");
//		/addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\LogoForExcel.png", 0, 0, 14, 17, "");
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 13, 18));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(1, 1, 13, 18), "");
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 13, 18));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(2, 2, 13, 18), "");
		sheet.addMergedRegion(new CellRangeAddress(3, 3, 13, 18));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(3, 3, 13, 18), "");
		
		
		
		sheet.addMergedRegion(new CellRangeAddress(4,11, 0, 0));
		
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4,11, 0, 0), "normalHeadding");
		normalHeadder1.getCell(0).setCellValue("Sl.No");
		sheet.addMergedRegion(new CellRangeAddress(4,11, 1, 1));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4,11, 1, 1), "normalHeadding");
		normalHeadder1.getCell(1).setCellValue("Location");
		sheet.addMergedRegion(new CellRangeAddress(4,11, 2, 2));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4,11, 2, 2), "normalHeadding");
		normalHeadder1.getCell(2).setCellValue("Equipment No");
		sheet.addMergedRegion(new CellRangeAddress(4,11, 3, 3));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4,11, 3, 3), "normalHeadding");
		normalHeadder1.getCell(3).setCellValue("Equipment Name");
		sheet.addMergedRegion(new CellRangeAddress(4,11, 4, 4));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4,11, 4, 4), "normalHeadding");
		normalHeadder1.getCell(4).setCellValue("Equipment Category");
		sheet.addMergedRegion(new CellRangeAddress(4,11, 5, 5));
		setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4,11, 5, 5), "normalHeadding");
		normalHeadder1.getCell(5).setCellValue("Equipment Class");
		
		normalHeadder1.createCell(6).setCellValue("Month");normalHeadder1.getCell(6).setCellStyle(styleSC);
		normalHeadder1.createCell(7).setCellValue("Apr");normalHeadder1.getCell(7).setCellStyle(styleLY);
		normalHeadder1.createCell(8).setCellValue("May");normalHeadder1.getCell(8).setCellStyle(styleLY);
		normalHeadder1.createCell(9).setCellValue("Jun");normalHeadder1.getCell(9).setCellStyle(styleLY);
		normalHeadder1.createCell(10).setCellValue("Jul");normalHeadder1.getCell(10).setCellStyle(styleLY);
		normalHeadder1.createCell(11).setCellValue("Aug");normalHeadder1.getCell(11).setCellStyle(styleLY);
		normalHeadder1.createCell(12).setCellValue("Sep");normalHeadder1.getCell(12).setCellStyle(styleLY);
		normalHeadder1.createCell(13).setCellValue("Oct");normalHeadder1.getCell(13).setCellStyle(styleLY);
		normalHeadder1.createCell(14).setCellValue("Nov");normalHeadder1.getCell(14).setCellStyle(styleLY);
		normalHeadder1.createCell(15).setCellValue("Dec");normalHeadder1.getCell(15).setCellStyle(styleLY);
		normalHeadder1.createCell(16).setCellValue("Jan");normalHeadder1.getCell(16).setCellStyle(styleLY);
		normalHeadder1.createCell(17).setCellValue("Feb");normalHeadder1.getCell(17).setCellStyle(styleLY);
		normalHeadder1.createCell(18).setCellValue("Mar");normalHeadder1.getCell(18).setCellStyle(styleLY);
		
		normalHeadder2.createCell(6).setCellValue("Days");normalHeadder2.getCell(6).setCellStyle(styleSC);
		normalHeadder2.createCell(7).setCellValue(daysPerMonthInFinancialYear[0]);normalHeadder2.getCell(7).setCellStyle(styleDY);
		normalHeadder2.createCell(8).setCellValue(daysPerMonthInFinancialYear[1]);normalHeadder2.getCell(8).setCellStyle(styleDY);
		normalHeadder2.createCell(9).setCellValue(daysPerMonthInFinancialYear[2]);normalHeadder2.getCell(9).setCellStyle(styleDY);
		normalHeadder2.createCell(10).setCellValue(daysPerMonthInFinancialYear[3]);normalHeadder2.getCell(10).setCellStyle(styleDY);
		normalHeadder2.createCell(11).setCellValue(daysPerMonthInFinancialYear[4]);normalHeadder2.getCell(11).setCellStyle(styleDY);
		normalHeadder2.createCell(12).setCellValue(daysPerMonthInFinancialYear[5]);normalHeadder2.getCell(12).setCellStyle(styleDY);
		normalHeadder2.createCell(13).setCellValue(daysPerMonthInFinancialYear[6]);normalHeadder2.getCell(13).setCellStyle(styleDY);
		normalHeadder2.createCell(14).setCellValue(daysPerMonthInFinancialYear[7]);normalHeadder2.getCell(14).setCellStyle(styleDY);
		normalHeadder2.createCell(15).setCellValue(daysPerMonthInFinancialYear[8]);normalHeadder2.getCell(15).setCellStyle(styleDY);
		normalHeadder2.createCell(16).setCellValue(daysPerMonthInFinancialYear[9]);normalHeadder2.getCell(16).setCellStyle(styleDY);
		normalHeadder2.createCell(17).setCellValue(daysPerMonthInFinancialYear[10]);normalHeadder2.getCell(17).setCellStyle(styleDY);
		normalHeadder2.createCell(18).setCellValue(daysPerMonthInFinancialYear[11]);normalHeadder2.getCell(18).setCellStyle(styleDY);
		
		normalHeadder3.createCell(6).setCellValue("");normalHeadder3.getCell(6).setCellStyle(styleSC);
		normalHeadder3.createCell(7).setCellValue("");normalHeadder3.getCell(7).setCellStyle(styleDY);
		normalHeadder3.createCell(8).setCellValue("");normalHeadder3.getCell(8).setCellStyle(styleDY);
		normalHeadder3.createCell(9).setCellValue("");normalHeadder3.getCell(9).setCellStyle(styleDY);
		normalHeadder3.createCell(10).setCellValue("");normalHeadder3.getCell(10).setCellStyle(styleDY);
		normalHeadder3.createCell(11).setCellValue("");normalHeadder3.getCell(11).setCellStyle(styleDY);
		normalHeadder3.createCell(12).setCellValue("");normalHeadder3.getCell(12).setCellStyle(styleDY);
		normalHeadder3.createCell(13).setCellValue("");normalHeadder3.getCell(13).setCellStyle(styleDY);
		normalHeadder3.createCell(14).setCellValue("");normalHeadder3.getCell(14).setCellStyle(styleDY);
		normalHeadder3.createCell(15).setCellValue("");normalHeadder3.getCell(15).setCellStyle(styleDY);
		normalHeadder3.createCell(16).setCellValue("");normalHeadder3.getCell(16).setCellStyle(styleDY);
		normalHeadder3.createCell(17).setCellValue("");normalHeadder3.getCell(17).setCellStyle(styleDY);
		normalHeadder3.createCell(18).setCellValue("");normalHeadder3.getCell(18).setCellStyle(styleDY);
		
		normalHeadder4.createCell(6).setCellValue("");normalHeadder4.getCell(6).setCellStyle(styleSC);
		normalHeadder4.createCell(7).setCellValue("");normalHeadder4.getCell(7).setCellStyle(styleDY);
		normalHeadder4.createCell(8).setCellValue("");normalHeadder4.getCell(8).setCellStyle(styleDY);
		normalHeadder4.createCell(9).setCellValue("");normalHeadder4.getCell(9).setCellStyle(styleDY);
		normalHeadder4.createCell(10).setCellValue("");normalHeadder4.getCell(10).setCellStyle(styleDY);
		normalHeadder4.createCell(11).setCellValue("");normalHeadder4.getCell(11).setCellStyle(styleDY);
		normalHeadder4.createCell(12).setCellValue("");normalHeadder4.getCell(12).setCellStyle(styleDY);
		normalHeadder4.createCell(13).setCellValue("");normalHeadder4.getCell(13).setCellStyle(styleDY);
		normalHeadder4.createCell(14).setCellValue("");normalHeadder4.getCell(14).setCellStyle(styleDY);
		normalHeadder4.createCell(15).setCellValue("");normalHeadder4.getCell(15).setCellStyle(styleDY);
		normalHeadder4.createCell(16).setCellValue("");normalHeadder4.getCell(16).setCellStyle(styleDY);
		normalHeadder4.createCell(17).setCellValue("");normalHeadder4.getCell(17).setCellStyle(styleDY);
		normalHeadder4.createCell(18).setCellValue("");normalHeadder4.getCell(18).setCellStyle(styleDY);
		
		normalHeadder5.createCell(6).setCellValue("");normalHeadder5.getCell(6).setCellStyle(styleSC);
		normalHeadder5.createCell(7).setCellValue("");normalHeadder5.getCell(7).setCellStyle(styleDY);
		normalHeadder5.createCell(8).setCellValue("");normalHeadder5.getCell(8).setCellStyle(styleDY);
		normalHeadder5.createCell(9).setCellValue("");normalHeadder5.getCell(9).setCellStyle(styleDY);
		normalHeadder5.createCell(10).setCellValue("");normalHeadder5.getCell(10).setCellStyle(styleDY);
		normalHeadder5.createCell(11).setCellValue("");normalHeadder5.getCell(11).setCellStyle(styleDY);
		normalHeadder5.createCell(12).setCellValue("");normalHeadder5.getCell(12).setCellStyle(styleDY);
		normalHeadder5.createCell(13).setCellValue("");normalHeadder5.getCell(13).setCellStyle(styleDY);
		normalHeadder5.createCell(14).setCellValue("");normalHeadder5.getCell(14).setCellStyle(styleDY);
		normalHeadder5.createCell(15).setCellValue("");normalHeadder5.getCell(15).setCellStyle(styleDY);
		normalHeadder5.createCell(16).setCellValue("");normalHeadder5.getCell(16).setCellStyle(styleDY);
		normalHeadder5.createCell(17).setCellValue("");normalHeadder5.getCell(17).setCellStyle(styleDY);
		normalHeadder5.createCell(18).setCellValue("");normalHeadder5.getCell(18).setCellStyle(styleDY);
		
		normalHeadder6.createCell(6).setCellValue("");normalHeadder6.getCell(6).setCellStyle(styleSC);
		normalHeadder6.createCell(7).setCellValue("");normalHeadder6.getCell(7).setCellStyle(styleDY);
		normalHeadder6.createCell(8).setCellValue("");normalHeadder6.getCell(8).setCellStyle(styleDY);
		normalHeadder6.createCell(9).setCellValue("");normalHeadder6.getCell(9).setCellStyle(styleDY);
		normalHeadder6.createCell(10).setCellValue("");normalHeadder6.getCell(10).setCellStyle(styleDY);
		normalHeadder6.createCell(11).setCellValue("");normalHeadder6.getCell(11).setCellStyle(styleDY);
		normalHeadder6.createCell(12).setCellValue("");normalHeadder6.getCell(12).setCellStyle(styleDY);
		normalHeadder6.createCell(13).setCellValue("");normalHeadder6.getCell(13).setCellStyle(styleDY);
		normalHeadder6.createCell(14).setCellValue("");normalHeadder6.getCell(14).setCellStyle(styleDY);
		normalHeadder6.createCell(15).setCellValue("");normalHeadder6.getCell(15).setCellStyle(styleDY);
		normalHeadder6.createCell(16).setCellValue("");normalHeadder6.getCell(16).setCellStyle(styleDY);
		normalHeadder6.createCell(17).setCellValue("");normalHeadder6.getCell(17).setCellStyle(styleDY);
		normalHeadder6.createCell(18).setCellValue("");normalHeadder6.getCell(18).setCellStyle(styleDY);
		
		normalHeadder7.createCell(6).setCellValue("");normalHeadder7.getCell(6).setCellStyle(styleSC);
		normalHeadder7.createCell(7).setCellValue("");normalHeadder7.getCell(7).setCellStyle(styleDY);
		normalHeadder7.createCell(8).setCellValue("");normalHeadder7.getCell(8).setCellStyle(styleDY);
		normalHeadder7.createCell(9).setCellValue("");normalHeadder7.getCell(9).setCellStyle(styleDY);
		normalHeadder7.createCell(10).setCellValue("");normalHeadder7.getCell(10).setCellStyle(styleDY);
		normalHeadder7.createCell(11).setCellValue("");normalHeadder7.getCell(11).setCellStyle(styleDY);
		normalHeadder7.createCell(12).setCellValue("");normalHeadder7.getCell(12).setCellStyle(styleDY);
		normalHeadder7.createCell(13).setCellValue("");normalHeadder7.getCell(13).setCellStyle(styleDY);
		normalHeadder7.createCell(14).setCellValue("");normalHeadder7.getCell(14).setCellStyle(styleDY);
		normalHeadder7.createCell(15).setCellValue("");normalHeadder7.getCell(15).setCellStyle(styleDY);
		normalHeadder7.createCell(16).setCellValue("");normalHeadder7.getCell(16).setCellStyle(styleDY);
		normalHeadder7.createCell(17).setCellValue("");normalHeadder7.getCell(17).setCellStyle(styleDY);
		normalHeadder7.createCell(18).setCellValue("");normalHeadder7.getCell(18).setCellStyle(styleDY);
		
		normalHeadder8.createCell(6).setCellValue("");normalHeadder8.getCell(6).setCellStyle(styleSC);
		normalHeadder8.createCell(7).setCellValue("");normalHeadder8.getCell(7).setCellStyle(styleDY);
		normalHeadder8.createCell(8).setCellValue("");normalHeadder8.getCell(8).setCellStyle(styleDY);
		normalHeadder8.createCell(9).setCellValue("");normalHeadder8.getCell(9).setCellStyle(styleDY);
		normalHeadder8.createCell(10).setCellValue("");normalHeadder8.getCell(10).setCellStyle(styleDY);
		normalHeadder8.createCell(11).setCellValue("");normalHeadder8.getCell(11).setCellStyle(styleDY);
		normalHeadder8.createCell(12).setCellValue("");normalHeadder8.getCell(12).setCellStyle(styleDY);
		normalHeadder8.createCell(13).setCellValue("");normalHeadder8.getCell(13).setCellStyle(styleDY);
		normalHeadder8.createCell(14).setCellValue("");normalHeadder8.getCell(14).setCellStyle(styleDY);
		normalHeadder8.createCell(15).setCellValue("");normalHeadder8.getCell(15).setCellStyle(styleDY);
		normalHeadder8.createCell(16).setCellValue("");normalHeadder8.getCell(16).setCellStyle(styleDY);
		normalHeadder8.createCell(17).setCellValue("");normalHeadder8.getCell(17).setCellStyle(styleDY);
		normalHeadder8.createCell(18).setCellValue("");normalHeadder8.getCell(18).setCellStyle(styleDY);
		
		
		 String[] staticValues = {"Pune ", "Nashik ", "Chennai "};
		 CellRangeAddressList addressList = new CellRangeAddressList(12, 13, 1, 4);
		XSSFRow RowDD = sheet.createRow(12);
		XSSFCell CellDD = RowDD.createCell(1);
		CellDD.setCellStyle(styleDY);
		 DataValidationHelper validationHelper = sheet.getDataValidationHelper();
	        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(staticValues);
	        DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
	        dataValidation.setSuppressDropDownArrow(true);
	        dataValidation.setShowErrorBox(true);
	        // Add the data validation to the sheet
	        sheet.addValidationData(dataValidation);
		
	        XSSFRow Row1 = sheet.createRow(15);
			XSSFCell Cell1 = Row1.createCell(1);
			Cell1.setCellValue("hii this ");
			Cell1.setCellFormula("UPPER(A1)");
			Cell1.setCellStyle(styleb);
//			Cell1.setCellStyle(stylef);
			CellDD.setCellStyle(styleDY);
		    sheet.setColumnWidth(0, 1000);
			sheet.setColumnWidth(1, 5000);
			sheet.setColumnWidth(2, 4000);
			sheet.setColumnWidth(3, 7000);
			sheet.setColumnWidth(4, 5000);
			sheet.setColumnWidth(5, 5000);
			sheet.setColumnWidth(6, 5000);
			sheet.setColumnWidth(7, 2000);
			sheet.setColumnWidth(8, 2000);
			sheet.setColumnWidth(9, 2000);
			sheet.setColumnWidth(10, 2000);
			sheet.setColumnWidth(11, 2000);
			sheet.setColumnWidth(12, 2000);
			sheet.setColumnWidth(13, 2000);
			sheet.setColumnWidth(14, 2000);
			sheet.setColumnWidth(15, 2000);
			sheet.setColumnWidth(16, 2000);
			sheet.setColumnWidth(17, 2000);
			sheet.setColumnWidth(18, 2000);

		FileOutputStream fout = new FileOutputStream(filePath);
		workbook.write(fout);
		workbook.close();
	}
	public static void addImage(XSSFWorkbook workbook, XSSFSheet sheet, String reportImage, int startrow, int endrow,
		int startCol, int endcolo, String cellTYpe) throws Exception {
		InputStream inputStream = new FileInputStream(reportImage);
		byte[] imageBytes = IOUtils.toByteArray(inputStream);
		int pictureIdx = workbook.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_PNG);
		CreationHelper creationHelper = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = creationHelper.createClientAnchor();
		anchor.setAnchorType(AnchorType.DONT_MOVE_DO_RESIZE);
		anchor.setCol1(startCol);
		anchor.setCol2(endcolo);
		anchor.setRow1(startrow);
		anchor.setRow2(endrow);
		Picture picture = drawing.createPicture(anchor, pictureIdx);
        picture.resize(1.001);
		int borderOffset = 10;
		anchor.setDx1(borderOffset);
		anchor.setDx2((int) (anchor.getDx1() + picture.getImageDimension().getWidth()));
		anchor.setDy1(borderOffset);
		anchor.setDy2((int) (anchor.getDy1() + picture.getImageDimension().getHeight()));

	}
	public static void setBordersToMergedCell(XSSFWorkbook workbook, XSSFSheet sheet, CellRangeAddress rangeAddress,
			String fonts) {
		// Iterate over each cell within the merged range
		for (int rowNum = rangeAddress.getFirstRow(); rowNum <= rangeAddress.getLastRow(); rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			if (row == null) {
				row = sheet.createRow(rowNum);
			}
			for (int colNum = rangeAddress.getFirstColumn(); colNum <= rangeAddress.getLastColumn(); colNum++) {
				XSSFCell cell = row.getCell(colNum);
				if (cell == null) {
					cell = row.createCell(colNum);
				}
				// Set border for each cell
				setBorderToCell(cell, workbook, fonts);
			}
		}
	}
	public static void setBorderToCell(XSSFCell cell, XSSFWorkbook workbook, String fonts) {

		XSSFColor blueColor = new XSSFColor(new byte[] { 0, 0, (byte) 255 });

		XSSFFont font = workbook.createFont();
		font.setBold(true);
		XSSFFont font1 = workbook.createFont();
		font1.setBold(true);
		font1.setFontHeight(15);
		
		XSSFCellStyle style1 = workbook.createCellStyle();
		style1.setWrapText(true);
		
		//changeddeepak
		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setWrapText(true);
		style2.setVerticalAlignment(VerticalAlignment.TOP);
		style2.setBorderTop(BorderStyle.THIN);
		style2.setBorderBottom(BorderStyle.THIN);
		style2.setBorderLeft(BorderStyle.THIN);
		style2.setBorderRight(BorderStyle.THIN);
		
		XSSFCellStyle style3 = workbook.createCellStyle();
		style3.setWrapText(true);
		style3.setVerticalAlignment(VerticalAlignment.CENTER);
		style3.setAlignment(HorizontalAlignment.CENTER);
		style3.setBorderTop(BorderStyle.THIN);
		style3.setBorderBottom(BorderStyle.THIN);
		style3.setBorderLeft(BorderStyle.THIN);
		style3.setBorderRight(BorderStyle.THIN);
		style3.setFont(font1);
		
		byte redSC = (byte) 250 ;
		byte greenSC = (byte) 215 ;
		byte blueSC = (byte) 200;
		XSSFColor SkinColor = new XSSFColor(new byte[] { redSC, greenSC, blueSC });
		
		XSSFCellStyle styleSC = (XSSFCellStyle) workbook.createCellStyle();
		styleSC.setBorderBottom(BorderStyle.THIN);
		styleSC.setBorderTop(BorderStyle.THIN);
		styleSC.setBorderRight(BorderStyle.THIN);
		styleSC.setBorderLeft(BorderStyle.THIN);
		styleSC.setWrapText(true);
		styleSC.setAlignment(HorizontalAlignment.CENTER);
		styleSC.setVerticalAlignment(VerticalAlignment.TOP);
		styleSC.setFillForegroundColor(SkinColor);
		styleSC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleSC.setFont(font);
		
		XSSFCellStyle style = workbook.createCellStyle();
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setWrapText(true);
		if (!fonts.contentEquals("")) {
			if (fonts.equals("fontBlue")) {
				font.setColor(XSSFColor.toXSSFColor(blueColor));
				style.setVerticalAlignment(VerticalAlignment.CENTER);
				style.setAlignment(HorizontalAlignment.CENTER);
			}

			style.setFont(font);
		}
		if(fonts.equals("headding")) {
			cell.setCellStyle(style3);
		}
		else if(fonts.equals("normalHeadding")) {
			cell.setCellStyle(styleSC);
		}
		else {
			cell.setCellStyle(style);
		}
		
	}
	public static int[] getDaysPerMonthInFinancialYear() {
        int startYear = LocalDate.now().getMonthValue() < 4 ? LocalDate.now().getYear() - 1 : LocalDate.now().getYear();
        int endYear = startYear + 1;
        int[] daysPerMonth = new int[12];
        int index = 0;
        for (int month = 4; month <= 12; month++) {
            YearMonth yearMonth = YearMonth.of(startYear, month);
            daysPerMonth[index++] = yearMonth.lengthOfMonth();
        }

        for (int month = 1; month <= 3; month++) {
            YearMonth yearMonth = YearMonth.of(endYear, month);
            daysPerMonth[index++] = yearMonth.lengthOfMonth();
        }
        return daysPerMonth;
    }
}

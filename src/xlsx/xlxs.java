package xlsx;

import java.awt.Graphics2D;
import java.awt.font.FontRenderContext;
import java.awt.font.TextLayout;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlxs {
	public static void main(String[] args) throws IOException {
		String baseString = "";

//		String Logo = mahindraLogo;
		String ECN_NO = "";
		String PAB_NO = "";
		String SAP_UPDATION = "";
		String Model = "";
		String Department = "";
		String ProblemNo = "";
		String CCRNo = "";
		String Source = "";
		String Severity = "";
		String RootCause = "";
		String ContainmentAction = "";
		String PermanentAction = "";
		ArrayList<String> PCACutOffNo = new ArrayList<String>();
		ArrayList<String> ICACutOffNo = new ArrayList<String>();
		ArrayList<String> PCADate = new ArrayList<String>();
		ArrayList<String> ICADate = new ArrayList<String>();
		RichTextString richText = null;
		String ClosureDate = "";
		String PCACutOffNo1 = "";
		String ICACutOffNo1 = "";
		String PCADate1 = "";
		String ICADate1 = "";
		String OBSERVATIONANALYSIS = "testestingtestingtestingtestingtestingtestestingtestingtestingtestingtestingtestingtestingtestingtestingtestingtestingtestingtingtestingtestingtestingtestingtestingtestingtestingting";
		String CustomerProtectAction = "";
		String BEFORE = "";
		String Trend = "";
		String After = "";
		String Plant = "";
		String ReportedDate = "";
		String ConcernDescription = "";
		String Remarks = "";
		String PDTCOE = "";
		String ConcernOwner = "";
		String PDTPlatformLead = "";
		String PDTHEAD = "";
		String PDTCOESignature = "";
		String ConcernOwnerSignature = "";
		String PDTPlatformLeadSignature = "";
		String PDTHEADSignature = "";
		String Pvorcv = "";
		String Version = "";
		String BusinessDivision = "";
		try {
			String filePath = "C:\\Users\\NICHEBIT\\Desktop\\test\\Testing31.xlsx";
			XSSFWorkbook workbook = new XSSFWorkbook();

			XSSFSheet sheet = workbook.createSheet();

			// Custom color RGB values
			byte red2 = (byte) 255;
			byte green2 = (byte) 229;
			byte blue2 = (byte) 152;

			XSSFColor ecuHeaderColorYell = new XSSFColor(new byte[] { red2, green2, blue2 });
			XSSFColor blueColor = new XSSFColor(new byte[] { 0, 0, (byte) 255 });

			XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();

			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);

			XSSFFont font = workbook.createFont();
			font.setFontHeight(20);
			font.setBold(true);

			XSSFFont font1 = workbook.createFont();
			font1.setBold(true);

			
			XSSFFont font2 = workbook.createFont();
			font1.setBold(true);
			XSSFCellStyle style1 = (XSSFCellStyle) workbook.createCellStyle();

			style1.setBorderBottom(BorderStyle.THIN);
			style1.setBorderTop(BorderStyle.THIN);
			style1.setBorderRight(BorderStyle.THIN);
			style1.setBorderLeft(BorderStyle.THIN);
			style1.setWrapText(true);
			style1.setAlignment(HorizontalAlignment.CENTER);
			style1.setFont(font1);

			XSSFCellStyle style3 = (XSSFCellStyle) workbook.createCellStyle();
			style1.setFont(font2);
			style3.setAlignment(HorizontalAlignment.LEFT);
			style3.setWrapText(true);
			
			XSSFCellStyle styleVM = (XSSFCellStyle) workbook.createCellStyle();

			styleVM.setBorderBottom(BorderStyle.THIN);
			styleVM.setBorderTop(BorderStyle.THIN);
			styleVM.setBorderRight(BorderStyle.THIN);
			styleVM.setBorderLeft(BorderStyle.THIN);
			styleVM.setVerticalAlignment(VerticalAlignment.CENTER);
			styleVM.setWrapText(true);
			
			
			
			
			

			XSSFCellStyle style2 = (XSSFCellStyle) workbook.createCellStyle();

			style2.setBorderBottom(BorderStyle.THIN);
			style2.setBorderTop(BorderStyle.THIN);
			style2.setBorderRight(BorderStyle.THIN);
			style2.setBorderLeft(BorderStyle.THIN);
			style2.setWrapText(true);
			style2.setAlignment(HorizontalAlignment.CENTER);
			style2.setVerticalAlignment(VerticalAlignment.CENTER);
			style2.setFont(font);

			XSSFCellStyle styleForRemoveRightBorder = (XSSFCellStyle) workbook.createCellStyle();
			styleForRemoveRightBorder.setBorderBottom(BorderStyle.THIN);
			styleForRemoveRightBorder.setBorderTop(BorderStyle.THIN);
			styleForRemoveRightBorder.setBorderLeft(BorderStyle.THIN);

			XSSFCellStyle styleForRemoveRightBorderb = (XSSFCellStyle) workbook.createCellStyle();
			styleForRemoveRightBorderb.setBorderBottom(BorderStyle.THIN);
			styleForRemoveRightBorderb.setBorderTop(BorderStyle.THIN);
			styleForRemoveRightBorderb.setBorderLeft(BorderStyle.THIN);
			styleForRemoveRightBorderb.setFont(font1);

			XSSFCellStyle styleForRemoveLeftBorder = (XSSFCellStyle) workbook.createCellStyle();
			styleForRemoveLeftBorder.setBorderBottom(BorderStyle.THIN);
			styleForRemoveLeftBorder.setBorderTop(BorderStyle.THIN);
			styleForRemoveLeftBorder.setBorderRight(BorderStyle.THIN);
			XSSFCellStyle styleForRemoveLeftBorderb = (XSSFCellStyle) workbook.createCellStyle();
			styleForRemoveLeftBorderb.setBorderBottom(BorderStyle.THIN);
			styleForRemoveLeftBorderb.setBorderTop(BorderStyle.THIN);
			styleForRemoveLeftBorderb.setBorderRight(BorderStyle.THIN);
			styleForRemoveLeftBorderb.setFont(font1);

			XSSFCellStyle styleleft = (XSSFCellStyle) workbook.createCellStyle();
			styleleft.setBorderLeft(BorderStyle.THIN);

			XSSFCellStyle styleright = (XSSFCellStyle) workbook.createCellStyle();
			styleright.setBorderRight(BorderStyle.THIN);

			XSSFCellStyle stylebottom = (XSSFCellStyle) workbook.createCellStyle();
			stylebottom.setBorderBottom(BorderStyle.THIN);

			XSSFCellStyle styleleftbottom = (XSSFCellStyle) workbook.createCellStyle();
			styleleftbottom.setBorderLeft(BorderStyle.THIN);
			styleleftbottom.setBorderBottom(BorderStyle.THIN);
			styleleftbottom.setAlignment(HorizontalAlignment.CENTER);
			XSSFCellStyle stylerightbottom = (XSSFCellStyle) workbook.createCellStyle();
			stylerightbottom.setBorderRight(BorderStyle.THIN);
			stylerightbottom.setBorderBottom(BorderStyle.THIN);

			XSSFCellStyle styleleftb = (XSSFCellStyle) workbook.createCellStyle();
			styleleftb.setBorderLeft(BorderStyle.THIN);
			styleleftb.setFont(font1);
			XSSFCellStyle stylerightb = (XSSFCellStyle) workbook.createCellStyle();
			stylerightb.setBorderRight(BorderStyle.THIN);
			stylerightb.setFont(font1);
			XSSFCellStyle stylebottomb = (XSSFCellStyle) workbook.createCellStyle();
			stylebottomb.setBorderBottom(BorderStyle.THIN);
			stylebottomb.setFont(font1);
			XSSFCellStyle styleleftbottomb = (XSSFCellStyle) workbook.createCellStyle();
			styleleftbottomb.setBorderLeft(BorderStyle.THIN);
			styleleftbottomb.setBorderBottom(BorderStyle.THIN);
			styleleftbottomb.setAlignment(HorizontalAlignment.CENTER);
			styleleftbottomb.setFont(font1);
			XSSFCellStyle stylerightbottomb = (XSSFCellStyle) workbook.createCellStyle();
			stylerightbottomb.setBorderRight(BorderStyle.THIN);
			stylerightbottomb.setBorderBottom(BorderStyle.THIN);
			stylerightbottomb.setFont(font1);

			XSSFCellStyle normalStyle = (XSSFCellStyle) workbook.createCellStyle();

			XSSFFont boldFont = workbook.createFont();
			boldFont.setBold(true);
			XSSFFont normalFont = workbook.createFont();

			XSSFRow row = null;
			XSSFRow r1 = sheet.createRow(0);
			XSSFRow r2 = sheet.createRow(1);
			XSSFRow r3 = sheet.createRow(2);

			sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 1));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(0, 2, 0, 1), "");
			r1.setHeight((short) 500);
			r2.setHeight((short) 500);
			r3.setHeight((short) 500);
			XSSFCell mergedCell = r1.createCell(0);
			setBorderToCell(mergedCell, workbook, "");
			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\LogoForExcel.png", 0, 3, 0, 2, "");
			mergedCell.setCellStyle(style);
			sheet.addMergedRegion(new CellRangeAddress(0, 2, 2, 5));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(0, 2, 2, 5), "");
			XSSFCell mergedCell2 = r1.createCell(2);
			mergedCell2.setCellValue("Concern Clouser Report");
			mergedCell2.setCellStyle(style2);

			XSSFCell cell = null;

			sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 7));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(0, 0, 6, 7), "");
			String plant = "Plant : ";
			richText = workbook.getCreationHelper().createRichTextString(plant + Plant);
			richText.applyFont(0, plant.length(), boldFont);
			if (Plant != "" && Plant != "") {
				richText.applyFont(plant.length() + 1, plant.length() + Plant.length(), normalFont);
			}
			cell = r1.getCell(6);
			cell.setCellValue(richText);
			
			cell.setCellStyle(styleVM);
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 7));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(1, 1, 6, 7), "");
			String department = "Department : ";
			richText = workbook.getCreationHelper().createRichTextString(department + Department);
			richText.applyFont(0, department.length(), boldFont);
			if (Department != "" && Department != "") {
				richText.applyFont(department.length() + 1, department.length() + Department.length(), normalFont);
			}
			cell = r2.getCell(6);
			cell.setCellValue(richText);
			cell.setCellStyle(styleVM);

			sheet.addMergedRegion(new CellRangeAddress(2, 2, 6, 7));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(2, 2, 6, 7), "");
			String model = "Model:";
			richText = workbook.getCreationHelper().createRichTextString(model + Model);
			richText.applyFont(0, model.length(), boldFont);
			if (Model != "" && Model != "") {
				richText.applyFont(model.length() + 1, model.length() + Model.length(), normalFont);
			}
			cell = r3.getCell(6);
			cell.setCellValue(richText);
			cell.setCellStyle(styleVM);

			// ROW3
			row = sheet.createRow(3);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(3, 3, 0, 1), "");
			String problemNo = "Problem No: ";
			richText = workbook.getCreationHelper().createRichTextString(problemNo + ProblemNo);
			richText.applyFont(0, problemNo.length(), boldFont);
			if (ProblemNo != "" && ProblemNo != "") {
				richText.applyFont(problemNo.length() + 1, problemNo.length() + ProblemNo.length(), normalFont);
			}
			cell = row.createCell(0);
			cell.setCellValue(richText);
			cell.setCellStyle(styleVM);
			sheet.addMergedRegion(new CellRangeAddress(3, 3, 2, 3));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(3, 3, 2, 3), "");
			String reportedDate = "Reported Date: ";
			richText = workbook.getCreationHelper().createRichTextString(reportedDate + ReportedDate);
			richText.applyFont(0, reportedDate.length(), boldFont);
			if (ReportedDate != "" && ReportedDate != "") {
				richText.applyFont(reportedDate.length() + 1, reportedDate.length() + ReportedDate.length(),
						normalFont);
			}
			cell = row.createCell(2);
			cell.setCellStyle(styleVM);
			cell.setCellValue(richText);

			sheet.addMergedRegion(new CellRangeAddress(3, 3, 4, 5));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(3, 3, 4, 5), "");
			String source = "Source: ";
			richText = workbook.getCreationHelper().createRichTextString(source + Source);
			richText.applyFont(0, source.length(), boldFont);
			if (Source != "" && Source != "") {
				richText.applyFont(source.length() + 1, source.length() + Source.length(), normalFont);
			}
			cell = row.createCell(4);
			cell.setCellStyle(styleVM);
			cell.setCellValue(richText);

			String severity = "Severity :";
			richText = workbook.getCreationHelper().createRichTextString(severity + Severity);
			richText.applyFont(0, severity.length(), boldFont);
			if (Severity != "" && Severity != "") {
				richText.applyFont(severity.length() + 1, severity.length() + Severity.length(), normalFont);
			}
			cell = row.createCell(6);
			cell.setCellStyle(styleVM);
			cell.setCellValue(richText);
			

			cell = row.createCell(7);
			String clouserDate = "Closure Date :";
			richText = workbook.getCreationHelper().createRichTextString(clouserDate + ClosureDate);
			richText.applyFont(0, clouserDate.length(), boldFont);
			if (ClosureDate != "" && ClosureDate != "") {
				richText.applyFont(clouserDate.length() + 1, clouserDate.length() + ClosureDate.length(), normalFont);
			}
			cell.setCellValue(richText);
			cell.setCellStyle(styleVM);

			// ROW4
			row = sheet.createRow(4);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 5));
			cell = row.createCell(0);
			String concernDescription = "Concern Description: ";
			richText = workbook.getCreationHelper().createRichTextString(concernDescription + ConcernDescription);
			richText.applyFont(0, concernDescription.length(), boldFont);
			if (ConcernDescription != "" && ConcernDescription != "") {
				richText.applyFont(concernDescription.length() + 1,
						concernDescription.length() + ConcernDescription.length(), normalFont);
			}
			cell.setCellValue(richText);
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4, 4, 0, 5), "font");
			sheet.addMergedRegion(new CellRangeAddress(4, 4, 6, 7));
			cell = row.createCell(6);

			String cCRNo = "CCR No: ";
			richText = workbook.getCreationHelper().createRichTextString(cCRNo + CCRNo);
			richText.applyFont(0, cCRNo.length(), boldFont);
			if (CCRNo != "" && CCRNo != "") {
				richText.applyFont(cCRNo.length() + 1, cCRNo.length() + CCRNo.length(), normalFont);
			}
			cell.setCellValue(richText);

			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(4, 4, 6, 7), "font");

			XSSFFont fontforOa = workbook.createFont();
			fontforOa.setBold(true);
			fontforOa.setColor(IndexedColors.BLUE.getIndex());
			XSSFCellStyle WRAPtEXT = (XSSFCellStyle) workbook.createCellStyle();
			WRAPtEXT.setWrapText(true);
			WRAPtEXT.setFont(font1);
			WRAPtEXT.setVerticalAlignment(VerticalAlignment.TOP);
			sheet.autoSizeColumn(0);
			

			XSSFCellStyle styleforOA = (XSSFCellStyle) workbook.createCellStyle();
			styleforOA.setFont(fontforOa);

			// ROW5
			row = sheet.createRow(5);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 3)); 	
			cell = row.createCell(0);
			cell.setCellValue("OBSERVATION & ANALYSIS : ");

			setBorderToCellLR(cell, workbook, "fontBlueAL");
			cell = row.createCell(1);
			setBorderToCellLR(cell, workbook, "fontBlueAL");
			cell = row.createCell(2);
			setBorderToCellLR(cell, workbook, "fontBlueAL");
			cell = row.createCell(3);
			setBorderToCellLR(cell, workbook, "fontBlueAL");
			sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 7));
			cell = row.createCell(4);
			cell.setCellValue("CORRECTIVE ACTIONS : ");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(5, 5, 4, 7), "fontBlueAL");

			// ROW6
			
			XSSFCellStyle  testingstyle=workbook.createCellStyle();
			testingstyle.setWrapText(true);
			testingstyle.setVerticalAlignment(VerticalAlignment.TOP);
		
			
			//changeddeepak
			
			row = sheet.createRow(6);
			
			sheet.addMergedRegion(new CellRangeAddress(6, 6, 0, 3));
			cell = row.createCell(0);
			 cell.setCellStyle(testingstyle);
			cell.setCellValue(OBSERVATIONANALYSIS);
			
			if(OBSERVATIONANALYSIS.length()>10) {
				if(OBSERVATIONANALYSIS.length()<85) {
					row.setHeightInPoints((short) ((OBSERVATIONANALYSIS.length() * 8 * 0.70) / 10));	
				}
				else {
					row.setHeightInPoints((short) ((OBSERVATIONANALYSIS.length() * 8 * 0.30) / 10));
				}
				
				
			}
			else {
				row.setHeightInPoints((short)25);
			}
			sheet.addMergedRegion(new CellRangeAddress(6, 6, 4, 7));
			cell = row.createCell(4);
			
			 
			 
			//changeddeepak
			String customerProtectAction = "Customer Protect Action :";
			richText = workbook.getCreationHelper().createRichTextString(customerProtectAction + CustomerProtectAction);
			richText.applyFont(0, customerProtectAction.length(), boldFont);
			if (CustomerProtectAction != "" && CustomerProtectAction != "") {
				richText.applyFont(customerProtectAction.length() + 1,
						customerProtectAction.length() + CustomerProtectAction.length(), normalFont);
			}
			if(richText.length()>OBSERVATIONANALYSIS.length()) {
				if(richText.length()>10) {
				if(richText.length()<85) {
					row.setHeightInPoints((short) ((richText.length() * 8 * 0.70) / 10));	
				}
				else {
					row.setHeightInPoints((short) ((richText.length() * 8 * 0.33) / 10));
				}
			}
				else {
					row.setHeightInPoints((short) 25);
				}
			}
			
			cell.setCellValue(richText);
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(6, 6, 4, 7), "fontTop");
		//	cell.setCellStyle(testingstyle);
			
           
			

			
			//ROW7
			XSSFCellStyle styleforrc = (XSSFCellStyle) workbook.createCellStyle();
			styleforrc.setFont(fontforOa);
			styleforrc.setWrapText(true);
			//changedDeepak
			styleforrc.setVerticalAlignment(VerticalAlignment.TOP);

			row = sheet.createRow(7);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 3));
			cell = row.createCell(0);
			cell.setCellValue("Root Cause :");
			cell.setCellStyle(styleforrc);
			sheet.addMergedRegion(new CellRangeAddress(7, 7, 4, 7));
			cell = row.createCell(4);
			//changeddeepak
			String containmentAction = "Containment Action : ";
			richText = workbook.getCreationHelper().createRichTextString(containmentAction + ContainmentAction);
			richText.applyFont(0, containmentAction.length(), boldFont);
			if (ContainmentAction != "" && ContainmentAction != "") {
				richText.applyFont(containmentAction.length() + 1, containmentAction.length() + ContainmentAction.length(),
						normalFont);
			}
			if(richText.length()<10) {
				if(richText.length()<85) {
					row.setHeightInPoints((short) ((richText.length() * 8 * 0.70) / 10));	
				}
				else {
					row.setHeightInPoints((short) ((richText.length() * 8 * 0.33) / 10));
				}
				
			}
			else {
				row.setHeightInPoints((short)25);
			}
			cell.setCellValue(richText);

			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(7, 7, 4, 7), "fontTop");
			
			
			//row9
			//changeddeepak
			row = sheet.createRow(8);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 3));
			cell = row.createCell(0);
			if(richText.length()<10) {
			if(richText.length()<85) {
				row.setHeightInPoints((short) ((RootCause.length() * 8 * 0.70) / 10));	
			}
			else {
				row.setHeightInPoints((short) ((RootCause.length() * 8 * 0.33) / 10));
			}
			}
			else {
				row.setHeightInPoints((short)25);
			}
			
			cell.setCellValue(RootCause);
			setBorderToCellLRb(cell, workbook, "");
			cell.setCellStyle(testingstyle);
			cell = row.createCell(1);
			setBorderToCellLRb(cell, workbook, "");
			cell.setCellStyle(testingstyle);
			cell = row.createCell(2);
			cell.setCellStyle(testingstyle);
			setBorderToCellLRb(cell, workbook, "");
			cell = row.createCell(3);
			setBorderToCellLRb(cell, workbook, "");
			cell.setCellStyle(testingstyle);
			

			sheet.addMergedRegion(new CellRangeAddress(8, 8, 4, 7));
			cell = row.createCell(4);
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(8, 8, 4, 7), "fontTop");
			String permanentAction = "Permanent Action : ";
			richText = workbook.getCreationHelper().createRichTextString(permanentAction + PermanentAction);
			richText.applyFont(0, permanentAction.length(), boldFont);
			if (PermanentAction != "" && PermanentAction != "") {
				richText.applyFont(permanentAction.length() + 1, permanentAction.length() + PermanentAction.length(),
						normalFont);
			}
			if(richText.length()>RootCause.length()) {
				if(richText.length()<10) {
				if(richText.length()<85) {
					row.setHeightInPoints((short) ((richText.length() * 8 * 0.70) / 10));	
				}
				else {
					row.setHeightInPoints((short) ((richText.length() * 8 * 0.33) / 10));
				}
			
			}
			}
			else {
				row.setHeightInPoints((short)25);
			}
			
			cell.setCellValue(richText);
			
			
			
			//row10//row11
			
			row = sheet.createRow(9);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(9, 9, 0, 7));
			cell = row.createCell(0);
			cell.setCellValue("Photograph/Nature Of Change	");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(9, 9, 0, 7), "fontBlue");
			row = sheet.createRow(10);
			row.setHeightInPoints(25);
			sheet.addMergedRegion(new CellRangeAddress(10, 10, 0, 3));
			cell = row.createCell(0);
			cell.setCellValue("BEFORE");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(10, 10, 0, 3), "fontBlue");

			sheet.addMergedRegion(new CellRangeAddress(10, 10, 4, 7));
			cell = row.createCell(4);
			cell.setCellValue("AFTER");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(10, 10, 4, 7), "fontBlue");
			
			//ROW12
			
			

			sheet.addMergedRegion(new CellRangeAddress(11, 13, 0, 3));

			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(11, 13, 0, 3), "");
			row = sheet.getRow(12);
			row.setHeight((short) 2000);
			BEFORE="C:\\Users\\NICHEBIT\\Downloads\\diagram (4).PNG";
			After=BEFORE;
			if (BEFORE != null && !BEFORE.isEmpty()) {
				addImage(workbook, sheet, BEFORE, 11, 14, 0, 3, "Before");
				addImageBorder(workbook, sheet, BEFORE, 11, 14, 0, 4, "Before");
			}
			sheet.addMergedRegion(new CellRangeAddress(11, 13, 4, 7));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(11, 13, 4, 7), "");
			if (After != null && !After.isEmpty()) {
				addImage(workbook, sheet, After, 11, 14, 4, 7, "Before");
				addImageBorder(workbook, sheet, BEFORE, 11, 14, 4, 8, "Before");
			}
			
			row = sheet.createRow(14);
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(14, 14, 0, 3));
			cell = row.createCell(0);
			cell.setCellValue("DFMEA RPN :");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(14, 14, 0, 3), "font");

			sheet.addMergedRegion(new CellRangeAddress(14, 14, 4, 7));
			cell = row.createCell(4);
			cell.setCellValue("DFMEA RPN3 :");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(14, 14, 4, 7), "font");

			// for Row 10

			int rowNo = 15;
			row = sheet.createRow(rowNo);
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			cell.setCellValue("Cut Off No.s :");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 0, 3), "fontBlueAL");

			row = sheet.createRow(++rowNo);
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			cell.setCellValue("Containment Action :");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 0, 3), "font");

			row = sheet.createRow(++rowNo);
			row.setHeightInPoints(20);
			cell = row.createCell(0);
			cell.setCellValue("Sr.No");
			cell.setCellStyle(style1);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 1, 2));
			cell = row.createCell(1);
			cell.setCellStyle(style1);
			cell.setCellValue("Cut Off No ");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 1, 2), "font");
			cell = row.createCell(3);
			cell.setCellStyle(style1);
			cell.setCellValue("Date");

			List<String> datas = new ArrayList<>();

			for (int i = 0; i < datas.size(); i++) {
				row = sheet.createRow(++rowNo);
				row.setHeightInPoints(20);
				cell = row.createCell(0);
				cell.setCellValue(i);
				cell.setCellStyle(style1);
				sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 1, 2));
				cell = row.createCell(1);

				cell.setCellValue(datas.get(i));
				setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 1, 2), "");
				cell = row.createCell(3);
				cell.setCellStyle(style1);
				cell.setCellValue(datas.get(i));

			}

			row = sheet.createRow(++rowNo);
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			cell.setCellValue("Permenent Action :");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 0, 3), "font");

			row = sheet.createRow(++rowNo);
			cell = row.createCell(0);
			row.setHeightInPoints(20);
			cell.setCellValue("Sr.No");
			cell.setCellStyle(style1);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 1, 2));
			cell = row.createCell(1);
			cell.setCellStyle(style1);
			cell.setCellValue("Cut Off No ");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 1, 2), "font");
			cell = row.createCell(3);
			cell.setCellStyle(style1);
			cell.setCellValue("Date");

			for (int i = 0; i < datas.size(); i++) {
				row = sheet.createRow(++rowNo);
				row.setHeightInPoints(20);
				cell = row.createCell(0);
				cell.setCellValue(i);
				cell.setCellStyle(style1);
				sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 1, 2));
				cell = row.createCell(1);

				cell.setCellValue(datas.get(i));
				setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 1, 2), "");
				cell = row.createCell(3);
				cell.setCellStyle(style1);
				cell.setCellValue(datas.get(i));

			}

			sheet.addMergedRegion(new CellRangeAddress(15, rowNo, 4, 7));
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(15, rowNo, 4, 7), "");
			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\diagram (4).PNG", 15, rowNo, 4, 7, "Before");
			addImageBorder(workbook, sheet, BEFORE, 15, rowNo+1, 4, 8, "Before");
//after nested tables
			row = sheet.createRow(++rowNo);
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			cell.setCellStyle(styleVM);
			cell.setCellValue("Documents Updated :");
			setBorderToCellLR(cell, workbook, "font");
			cell = row.createCell(1);
			setBorderToCellLR(cell, workbook, "font");
			cell = row.createCell(2);
			setBorderToCellLR(cell, workbook, "font");
			cell = row.createCell(3);
			setBorderToCellLR(cell, workbook, "font");
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 4, 7));
			cell = row.createCell(4);
			cell.setCellStyle(styleVM);
			cell.setCellValue("Sustenance Plan Updated :");
			setBorderToCellLR(cell, workbook, "font");
			cell = row.createCell(5);
			setBorderToCellLR(cell, workbook, "font");
			cell = row.createCell(6);
			setBorderToCellLR(cell, workbook, "font");
			cell = row.createCell(7);
			setBorderToCellLR(cell, workbook, "font");

			XSSFFont fontforr = workbook.createFont();
			fontforr.setBold(true);

			XSSFCellStyle styleforr = (XSSFCellStyle) workbook.createCellStyle();
			styleforr.setFont(fontforr);

			
			XSSFFont fontForcb = workbook.createFont();
			fontForcb.setBold(true);
			fontForcb.setFontHeightInPoints((short)20);
			XSSFCellStyle styleForcb = (XSSFCellStyle) workbook.createCellStyle();
			styleForcb.setFont(fontForcb);
//			styleForcb.setBorderBottom(BorderStyle.THIN);
//			styleForcb.setBorderTop(BorderStyle.THIN);
//			styleForcb.setBorderRight(BorderStyle.THIN);
//			styleForcb.setBorderLeft(BorderStyle.THIN);
			styleForcb.setVerticalAlignment(VerticalAlignment.BOTTOM);
			styleForcb.setAlignment(HorizontalAlignment.RIGHT);
			
			
			
			// checkoxData
			row = sheet.createRow(++rowNo);
			cell = row.createCell(0);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 0, 0, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 0, 1, "Before");
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
			cell = row.createCell(1);
			cell.setCellStyle(styleForcb);
			cell.setCellValue("POS / Process Shee");
			cell.setCellStyle(style3);

			cell = row.createCell(2);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 2, 2, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 2, 3, "Before");
			cell.setCellValue("☑");
			cell.setCellStyle(styleForcb);
			cell = row.createCell(3);
			cell.setCellValue("SOS/SOP,DCP");
			cell.setCellStyle(styleright);

			cell = row.createCell(4);
			cell.setCellValue("☑");
			cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 4, 4, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 4, 5, "Before");
//			cell.setCellStyle(styleleft);

			cell = row.createCell(5);
			cell.setCellValue("Self Check & Self Check Audit");
			cell.setCellStyle(style3);

			cell = row.createCell(6);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 6, 6, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 6, 7, "Before");

			cell = row.createCell(7);
			cell.setCellValue(" Supervisor Checklist");
			cell.setCellStyle(styleright);
			// 2nd row
			row = sheet.createRow(++rowNo);
			cell = row.createCell(0);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 0, 0, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 0, 1, "Before");
//			cell.setCellStyle(styleleft);

			cell = row.createCell(1);
			cell.setCellValue("Monitoring Plan Control Plan ");
			cell.setCellStyle(style3);

			cell = row.createCell(2);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 2, 2, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 2, 3, "Before");

			cell = row.createCell(3);
			cell.setCellValue("Checkman Checklist");
			cell.setCellStyle(styleright);

			cell = row.createCell(4);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 4, 4, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 4, 5, "Before");
//			cell.setCellStyle(styleleft);

			cell = row.createCell(5);
			cell.setCellValue("Checkman Checklist(FIXED PART)");
			cell.setCellStyle(style3);

			cell = row.createCell(6);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo, 6, 6, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 6, 7, "Before");

			cell = row.createCell(7);
			cell.setCellValue("JH Checklist");
			cell.setCellStyle(styleright);

			// checkoxData
			row = sheet.createRow(++rowNo);
			cell = row.createCell(0);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 0, 0, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 0, 1, "Before");
//			cell.setCellStyle(styleleft);

			cell = row.createCell(1);
			cell.setCellValue("OTHERS");
			cell.setCellStyle(style3);

			cell = row.createCell(2);
			cell.setCellValue("");

			cell = row.createCell(3);
			cell.setCellValue("");
			cell.setCellStyle(styleright);

			cell = row.createCell(4);
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo, 4, 4, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 4, 5, "Before");
//			cell.setCellStyle(styleleft);

			cell = row.createCell(5);
			cell.setCellValue("PM Checklist");
			cell.setCellStyle(style3);

			cell = row.createCell(6);
			
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\UncheckForEXcel.png", rowNo, rowNo, 6, 6, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 6, 7, "Before");

			cell = row.createCell(7);
			cell.setCellValue(" Process Audit");
			cell.setCellStyle(styleright);
			// 2nd row
			row = sheet.createRow(++rowNo);
			cell = row.createCell(0);
			cell.setCellValue("");
			cell.setCellStyle(styleleftbottom);

			cell = row.createCell(1);
			cell.setCellValue(" ");
			cell.setCellStyle(stylebottom);

			cell = row.createCell(2);
			cell.setCellValue("");
			cell.setCellStyle(stylebottom);

			cell = row.createCell(3);
			cell.setCellValue("");
			cell.setCellStyle(stylerightbottom);

			cell = row.createCell(4);
			
			cell.setCellValue("☑");cell.setCellStyle(styleForcb);
			//addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo, 4, 4, "check");
//			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Desktop\\CheckForExcel.png", rowNo, rowNo+1, 4, 5, "Before");
//			cell.setCellStyle(stylebottom);

			cell = row.createCell(5);
			cell.setCellValue("OTHERS");
			cell.setCellStyle(style3);

			cell = row.createCell(6);
			cell.setCellValue("");
			cell.setCellStyle(stylebottom);

			cell = row.createCell(7);
			cell.setCellValue("");
			cell.setCellStyle(stylerightbottom);

			
			
			
			
			// last row
			row = sheet.createRow(++rowNo);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			row.setHeightInPoints(20);
			
			cell = row.createCell(0);
			
			
			
			String remarks = "Remarks :";
			richText = workbook.getCreationHelper().createRichTextString(remarks + Remarks);
			richText.applyFont(0, remarks.length(), boldFont);
			if (Remarks != "" && Remarks != "") {
				richText.applyFont(remarks.length() + 1, remarks.length() + Remarks.length(), normalFont);
			}
			cell.setCellValue(richText);
			
			
			
			
			cell = row.createCell(1);
			
			cell = row.createCell(2);
			
			cell = row.createCell(3);
			
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 4, 7));
			cell = row.createCell(4);
			cell.setCellValue("Authorisation :");
			setBordersToMergedCell(workbook, sheet, new CellRangeAddress(rowNo, rowNo, 4, 7), "font");

			row = sheet.createRow(++rowNo);
			
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			
			
			String ECNNO = "ECA NO :";
			richText = workbook.getCreationHelper().createRichTextString(ECNNO + ECN_NO);
			richText.applyFont(0, ECNNO.length(), boldFont);
			if (ECN_NO != "" && ECN_NO != "") {
				richText.applyFont(ECNNO.length() + 1, ECNNO.length() + ECN_NO.length(), normalFont);
			}
			cell.setCellValue(richText);
			cell = row.createCell(4);
			cell.setCellStyle(style1);
			cell.setCellValue("Concern Owner");
			cell = row.createCell(5);
			cell.setCellValue("Platform Lead");
			cell.setCellStyle(style1);
			
			cell = row.createCell(6);
			cell.setCellValue("COE");
			cell.setCellStyle(style1);
			
			cell = row.createCell(7);
			cell.setCellValue("HEAD");
			cell.setCellStyle(style1);
			row = sheet.createRow(++rowNo);
			row.setHeightInPoints(20);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			String PABNO = "PAB NO :";
			richText = workbook.getCreationHelper().createRichTextString(PABNO + PAB_NO);
			richText.applyFont(0, PABNO.length(), boldFont);
			if (PAB_NO != "" && PAB_NO != "") {
				richText.applyFont(PABNO.length() + 1, PABNO.length() + PAB_NO.length(), normalFont);
			}
			cell.setCellValue(richText);
		
			cell = row.createCell(4);
			cell.setCellStyle(style1);
			cell.setCellValue("");
			cell = row.createCell(5);
			cell.setCellStyle(style1);
			cell.setCellValue("");
			cell = row.createCell(6);
			cell.setCellStyle(style1);
			cell.setCellValue("");
			cell = row.createCell(7);
			cell.setCellStyle(style1);
			cell.setCellValue("");

			row = sheet.createRow(++rowNo);
			row.setHeight((short)1300);

			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 0, 3));
			cell = row.createCell(0);
			setBorderToCellLRb(cell, workbook, "fontBlueAL");
			String sapUpdation = "SAP UPDATION :";
			richText = workbook.getCreationHelper().createRichTextString(sapUpdation + SAP_UPDATION);
			richText.applyFont(0, sapUpdation.length(), boldFont);
			if (SAP_UPDATION != "" && SAP_UPDATION != "") {
				richText.applyFont(sapUpdation.length() + 1, sapUpdation.length() + SAP_UPDATION.length(), normalFont);
			}
			cell.setCellValue(richText);
			setBorderToCellLRb(cell, workbook, "fontBlueAL");
			cell = row.createCell(1);
			setBorderToCellLRb(cell, workbook, "fontBlueAL");
			cell = row.createCell(2);
			setBorderToCellLRb(cell, workbook, "fontBlueAL");
			cell = row.createCell(3);
			setBorderToCellLRb(cell, workbook, "fontBlueAL");

			cell = row.createCell(4);
			cell.setCellStyle(style1);
			
			
			

			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo, 4, 4,
					"signature");

			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo+1,4, 5, "Before");
			cell = row.createCell(5);


			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo, 5, 5,"signature");
			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo+1, 5, 6, "Before");
			cell.setCellStyle(style1);
			cell = row.createCell(6);
//			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo, 6, 6,
//					"signature");

			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo, 6, 6,"signature");
			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo+1,6, 7, "Before");
			cell.setCellStyle(style1);

			cell = row.createCell(7);

			addImage(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo, 7, 7,"signature");
			addImageBorder(workbook, sheet, "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg", rowNo, rowNo+1,7, 8, "Before");
			 cell.setCellStyle(style1);

			row = sheet.createRow(36);
			row.setHeight((short) 1500);
			cell = row.createCell(7);
			addImageForCCR(workbook, sheet);

			XSSFCellStyle style111 = workbook.createCellStyle();
			Font font111 = workbook.createFont();
			font111.setFontName("Arial");
			font111.setFontHeightInPoints((short) 18);
			font111.setColor(IndexedColors.RED.getIndex());
			style111.setFont(font111);

			row = sheet.createRow(37);
			row.setHeight((short) 380);
			cell = row.createCell(1);

//String CCRNo="deepak";
			String name1 = "☑";
			richText = workbook.getCreationHelper().createRichTextString(name1 + "   " + CCRNo);
			richText.applyFont(0, name1.length(), font111);
			if (CCRNo != "" && CCRNo != "") {
				richText.applyFont(name1.length() + 1, name1.length() + CCRNo.length(), font1);
			}
			cell.setCellValue(richText);

			cell.setCellStyle(style111);
			cell = row.createCell(2);
			cell.setCellValue("☐");
			cell.setCellStyle(style111);
			
			//R&D For cell dynamic height
			
			
			
			
			
			
			row = sheet.createRow(38);
			row.setHeight((short)-1); 
			row.setHeightInPoints((short)-1);
			sheet.addMergedRegion(new CellRangeAddress(38, 38, 0, 3));
			cell = row.createCell(0);
			
	        cell.setCellStyle(testingstyle);
			cell.setCellValue("r");
			 sheet.autoSizeColumn(0);
			sheet.addMergedRegion(new CellRangeAddress(38, 38, 4, 7));
			
			
			String data="TestingTestTestingTestigTestiinTestingTTestingTestingTestTestingTestigTestiinTestingTTestingTestigTestiinTestingTTestingTestigTestiinTestingTigTestiinTestingTTestigTestiinTestingTTestingTestigTestiinTestingTigTestiinTestingT";
			int datalen=data.length();
			 System.out.println("length"+datalen);
			if(datalen<85)
			{
				row.setHeightInPoints((short) ((datalen * 8 * 0.75) / 10));
			}
			else {
				row.setHeightInPoints((short) ((datalen * 8 * 0.30) / 10));
			}
			cell = row.createCell(4);
			 cell.setCellStyle(testingstyle);
			 cell.setCellValue(data);
			 setBordersToMergedCell(workbook, sheet, new CellRangeAddress(38, 38, 4, 7), "testing");
			 
			 
			 row=sheet.getRow(38);
			 cell=row.getCell(0);
			 int colwidthinchars = (sheet.getColumnWidth(0) + sheet.getColumnWidth(1)+ sheet.getColumnWidth(2)+ sheet.getColumnWidth(3)) / 256;
			 System.out.println(colwidthinchars);
			 float cellHeight = getCellContentHeight(workbook, sheet, cell);
			 colwidthinchars = Math.round(colwidthinchars * 12f);
			System.out.println(colwidthinchars);
			
			
			 row=sheet.createRow(39);
			 sheet.addMergedRegion(new CellRangeAddress(39, 39, 0, 2));
			 setBordersToMergedCell(workbook, sheet, new CellRangeAddress(39, 39, 0, 2), "testing");
			 cell = row.createCell(0);
			row.setHeightInPoints((short) ((datalen * 8 * 0.30) / 10));
			 cell.setCellStyle(testingstyle);
			 cell.setCellValue(data);
			 
			 sheet.setColumnWidth(0, 3000);
			sheet.setColumnWidth(1, 7000);
			sheet.setColumnWidth(2, 3000);
			sheet.setColumnWidth(3, 7000);
			sheet.setColumnWidth(4, 5000);
			sheet.setColumnWidth(5, 5000);
			sheet.setColumnWidth(6, 5000);
			sheet.setColumnWidth(7, 5000);
		
			FileOutputStream fout = new FileOutputStream(filePath);
			workbook.write(fout);
			workbook.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	
	private static float getCellContentHeight(XSSFWorkbook workbook, XSSFSheet sheet, XSSFCell cell) {
		// TODO Auto-generated method stub
		return 0;
	}


	public static void addImg(XSSFWorkbook workbook, XSSFSheet sheet) throws Exception {
		FileInputStream fis = new FileInputStream("C:\\Users\\NICHEBIT\\Desktop\\si.png");
		byte[] imageBytes = IOUtils.toByteArray(fis);
		int pictureIndex = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
		CreationHelper helper = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();

		// Calculate cell dimensions in EMUs (English Metric Units)
		int cellWidth = (int) (sheet.getColumnWidthInPixels(5) - sheet.getColumnWidthInPixels(4));
		int cellHeight = (int) ((sheet.getRow(30).getHeightInPoints() / 72) * 96); // Convert points to pixels
		List<XSSFPictureData> allPictures = workbook.getAllPictures();
		System.out.println(allPictures.toString());
		// Calculate image dimensions in EMUs
//	        int imageWidth = workbook.getPictureData(pictureIndex).getImageDimension().width;
//	        int imageHeight = workbook.getPictureData(pictureIndex).getImageDimension().height;

		// Calculate x-axis and y-axis coordinates
//	        int x = (int) ((cellWidth - imageWidth) / 2 + sheet.getColumnWidthInPixels(4));
//	        int y = (cellHeight - imageHeight) / 2 + (int) (sheet.getRow(30).getHeightInPoints() / 72 * 96); // Convert points to pixels

		// Set the anchor to the cell range where you want to insert the image
		anchor.setCol1(4);
		anchor.setRow1(30);
		anchor.setCol2(5);
		anchor.setRow2(30);

		// Set x-axis and y-axis coordinates
//	        anchor.setDx1(x); // The x-coordinate within the first cell
//	        anchor.setDx2(x + imageWidth); // The x-coordinate within the last cell
//	        anchor.setDy1(y); // The y-coordinate within the first cell
//	        anchor.setDy2(y + imageHeight); // The y-coordinate within the last cell

		// Create the picture and anchor it to the cell
		Picture picture = drawing.createPicture(anchor, pictureIndex);

		// Resize the image to fit within the cell size
		picture.resize();
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
		if(fonts.equals("testing")) {
			cell.setCellStyle(style1);
		}
		else if(fonts.equals("fontTop")) {
			cell.setCellStyle(style2);
		}
		else {
			cell.setCellStyle(style);
		}
		
	}

	public static void setBorderToCellLR(XSSFCell cell, XSSFWorkbook workbook, String fonts) {

		XSSFColor blueColor = new XSSFColor(new byte[] { 0, 0, (byte) 255 });

		XSSFFont font = workbook.createFont();
		font.setBold(true);
		XSSFCellStyle style = workbook.createCellStyle();

		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		if (!fonts.contentEquals("")) {
			if (fonts.equals("fontBlue")) {
				font.setColor(XSSFColor.toXSSFColor(blueColor));
				style.setAlignment(HorizontalAlignment.CENTER);
			} else if (fonts.equals("fontBlueAL")) {
				style.setVerticalAlignment(VerticalAlignment.CENTER);
				font.setColor(XSSFColor.toXSSFColor(blueColor));
			}

			style.setFont(font);
		}
		cell.setCellStyle(style);
	}

	public static void setBorderToCellLRb(XSSFCell cell, XSSFWorkbook workbook, String fonts) {

		XSSFColor blueColor = new XSSFColor(new byte[] { 0, 0, (byte) 255 });

		XSSFFont font = workbook.createFont();
		font.setBold(true);
		XSSFCellStyle style = workbook.createCellStyle();
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		if (!fonts.contentEquals("")) {
			if (fonts.equals("fontBlue")) {
				font.setColor(XSSFColor.toXSSFColor(blueColor));
				style.setAlignment(HorizontalAlignment.CENTER);
			}

			style.setFont(font);
		}
		cell.setCellStyle(style);
	}

	public static void addImageBorder(XSSFWorkbook workbook, XSSFSheet sheet, String reportImage, int startrow, int endrow,
			int startCol, int endcolo, String cellTYpe) throws Exception {
		Drawing drawing = sheet.createDrawingPatriarch();
		 XSSFClientAnchor borderAnchor = new XSSFClientAnchor();
	        borderAnchor.setCol1(startCol);
	        borderAnchor.setRow1(startrow);
	        borderAnchor.setCol2(endcolo);
	        borderAnchor.setRow2(endrow);
    XSSFSimpleShape border = ((XSSFDrawing) drawing).createSimpleShape((XSSFClientAnchor) borderAnchor);
    border.setShapeType(ShapeTypes.RECT);
    border.setLineStyleColor(0, 0, 0); // RGB color for the border
    border.setLineWidth(1.0);
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
		// Set the anchor properties

		anchor.setCol1(startCol);
		anchor.setCol2(endcolo);
		anchor.setRow1(startrow);
		anchor.setRow2(endrow);

		// Create the picture and resize it
		Picture picture = drawing.createPicture(anchor, pictureIdx);

//		 XSSFClientAnchor borderAnchor = new XSSFClientAnchor();
//	        borderAnchor.setCol1(startCol);
//	        borderAnchor.setRow1(endcolo);
//	        borderAnchor.setCol2(endcolo);
//	        borderAnchor.setRow2(endrow);
//        XSSFSimpleShape border = ((XSSFDrawing) drawing).createSimpleShape((XSSFClientAnchor) borderAnchor);
//        border.setShapeType(ShapeTypes.RECT);
//        border.setLineStyleColor(0, 0, 0); // RGB color for the border
//        border.setLineWidth(1.0);

		if (cellTYpe.equals("check")) {
			picture.resize(1.001);
		} else if (cellTYpe.equals("Before")) {
			picture.resize(1);
		} else if (cellTYpe.equals("signature")) {
			picture.resize(1);
		} else if (cellTYpe.equals("si")) {
			picture.resize(1.0);
		}

		else {
			picture.resize(1.001);
		}
		int borderOffset = 10; // Adjust this value as needed
		anchor.setDx1(borderOffset);
		anchor.setDx2((int) (anchor.getDx1() + picture.getImageDimension().getWidth()));
		anchor.setDy1(borderOffset);
		anchor.setDy2((int) (anchor.getDy1() + picture.getImageDimension().getHeight()));

	}

	public static void addImgs(XSSFWorkbook workbook, XSSFSheet sheet, String ImageName, int StartingRow, int EndingRow,
			int StartingColoumn, int EndingColoumn) throws Exception {
		InputStream inputStream = new FileInputStream(ImageName);
		byte[] imageBytes = IOUtils.toByteArray(inputStream);
		inputStream.close();
		int cellWidthInPixels = 55; // Assuming the cell width is in pixels
		int cellHeightInPixels = 500; // Assuming the cell height is in pixels

		// Resize the image to fit cell dimensions
		byte[] resizedImageBytes = resizeImage(imageBytes, cellWidthInPixels, cellHeightInPixels);

		// Insert the resized image into the Excel sheet
		int pictureIndex = workbook.addPicture(resizedImageBytes, Workbook.PICTURE_TYPE_PNG);
		CreationHelper helper = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setCol1(StartingColoumn);
		anchor.setRow1(StartingRow);
		anchor.setCol2(EndingColoumn);
		anchor.setRow2(EndingRow);
		anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
		anchor.setDx1(1 * Units.EMU_PER_PIXEL);
		anchor.setDy1(1 * Units.EMU_PER_PIXEL);
		Picture picture = drawing.createPicture(anchor, pictureIndex);
		picture.resize();
	}

	private static byte[] resizeImage(byte[] originalImageBytes, int maxWidth, int maxHeight) throws IOException {
		// Convert bytes to an image
		InputStream in = new ByteArrayInputStream(originalImageBytes);
		BufferedImage originalImage = ImageIO.read(in);

		// Calculate the new dimensions while preserving aspect ratio
		int newWidth = originalImage.getWidth();
		int newHeight = originalImage.getHeight();
		double widthRatio = (double) maxWidth / originalImage.getWidth();
		double heightRatio = (double) maxHeight / originalImage.getHeight();
		double scaleFactor = Math.min(widthRatio, heightRatio);
		newWidth = (int) (originalImage.getWidth() * scaleFactor);
		newHeight = (int) (originalImage.getHeight() * scaleFactor);

		// Resize the image
		BufferedImage resizedImage = new BufferedImage(newWidth, newHeight, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2d = resizedImage.createGraphics();
		g2d.drawImage(originalImage, 0, 0, newWidth, newHeight, null);
		g2d.dispose();

		// Convert the resized image back to bytes
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ImageIO.write(resizedImage, "png", baos);
		baos.flush();
		byte[] resizedImageBytes = baos.toByteArray();
		baos.close();

		return resizedImageBytes;
	}

	public static void addImageForCCR(XSSFWorkbook workbook, XSSFSheet sheet) throws Exception {
		String picturePath = "C:\\Users\\NICHEBIT\\Downloads\\signature1.jpg";
		byte[] pictureBytes = org.apache.poi.util.IOUtils.toByteArray(new java.io.FileInputStream(picturePath));
		int pictureIdx = workbook.addPicture(pictureBytes, Workbook.PICTURE_TYPE_JPEG);

		CreationHelper helper = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setCol1(7); // Start column for the image
		anchor.setRow1(36); // Start row for the image
		anchor.setCol2(8); // End column for the image
		anchor.setRow2(37); // End row for the image (assuming the image will occupy two rows)

		Picture picture = drawing.createPicture(anchor, pictureIdx);
		picture.resize(1); // Resize the picture as needed

		// Apply border to the picture
		XSSFClientAnchor borderAnchor = new XSSFClientAnchor();
		borderAnchor.setCol1(7); // Start column for the border
		borderAnchor.setRow1(36); // Start row for the border
		borderAnchor.setCol2(8); // End column for the border
		borderAnchor.setRow2(37); // End row for the border (assuming the border will be around the image)

		XSSFSimpleShape border = ((XSSFDrawing) drawing).createSimpleShape((XSSFClientAnchor) borderAnchor);
		border.setShapeType(ShapeTypes.RECT);
		border.setLineStyleColor(0, 0, 0); // RGB color for the border
		border.setLineWidth(1.0); // Border width
	}
	
	

}

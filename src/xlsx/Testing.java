package xlsx;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Testing {

	public static void main(String args[]) throws Throwable {

		Workbook workbook = null;
		String Filepath = "C:\\Users\\NICHEBIT\\Downloads\\testingxlsxv.xlsx";
		FileInputStream fis = new FileInputStream(Filepath);

		if (Filepath.endsWith(".xlsx")) {
			workbook = new XSSFWorkbook(fis);

		} else if (Filepath.endsWith(".xls")) {

			workbook = new HSSFWorkbook(fis);
		} else {
			// throw the exception

		}

		String[] excelHeading = { "Cordys No", "Cancelled / Endorsement Proposal No", "UTR Number", "UTR Date",
				"UTR amount", "Customer Account No", "IFSC code" };
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> xlRows = sheet.iterator();
		int rowNo = 0;
		while (xlRows.hasNext()) {
			Row row = xlRows.next();
			rowNo++;
			double cordysNo;
			double CanEndPropNo;
			String UTRNumber;
			String UTRDate = "";
			double UTRamount;
			String CustomerAccountNo;
			String IFSCcode;
			Date dateUtr;
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

			if (rowNo == 1) {

				for (int i = 0; i < excelHeading.length; i++) {

					if (row.getCell(i).getStringCellValue().equals(excelHeading[i])) {

					} else {
						// through Exception
					}

				}
			}

			else {
				for (int j = 0; j < row.getLastCellNum(); j++) {

					// cell0
					if (j == 0) {
						System.out.println(row.getCell(j).getNumericCellValue());
						if (row.getCell(j) == null) {
							System.out.println("Empty Data");
							break;

						} else {
							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {

								cordysNo = row.getCell(j).getNumericCellValue();
								long cordysNoLo = (long) cordysNo;
								

							} else {
								// Failure case Data type is mismatch
								System.out.println("Data type is mismatch");
								break;

							}

						}

					}

					// cell1
					if (j == 1) {

						if (row.getCell(j) == null) {

							break;

						} else {

							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {

								CanEndPropNo = row.getCell(j).getNumericCellValue();
								long CanEndPropNoLo = (long) CanEndPropNo;

							} else {
								// Failure case Data type is mismatch
								break;

							}

						}

					}
					// cell2
					if (j == 2) {
						if (row.getCell(j) == null) {

							break;

						}

						else {

							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {
								UTRNumber = row.getCell(j).getStringCellValue();

							} else {
								// Failure case Data type is mismatch
								break;

							}
						}

					}
					// cell3

					if (j == 3) {

						if (row.getCell(j) == null) {

							break;

						} else {
							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC
									&& DateUtil.isCellDateFormatted(row.getCell(j))) {

								double numericValue = row.getCell(3).getNumericCellValue();
								java.util.Date dateValue = DateUtil.getJavaDate(numericValue);

								UTRDate = dateFormat.format(dateValue);
								java.util.Date utilDate = dateFormat.parse(UTRDate);
								java.sql.Date sqlDate = new java.sql.Date(utilDate.getTime());

								System.out.println(sqlDate);

							} else {
								// Failure case Data type is mismatch
								break;

							}

						}

					}
					if (j == 4) {

						if (row.getCell(j) == null) {

							break;

						} else {
							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {

								UTRamount = row.getCell(j).getNumericCellValue();
								long UTRamountLo = (long) UTRamount;

							} else {
								// Failure case Data type is mismatch
								break;

							}
						}

					}
					if (j == 5) {
						if (row.getCell(j) == null) {

							break;

						} else {
							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {

								CustomerAccountNo = row.getCell(j).getStringCellValue();

							} else {
								// Failure case Data type is mismatch
								break;

							}
						}

					}
					if (j == 6) {

						if (row.getCell(j) == null) {

							break;

						}

						else {
							if (row.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {

								IFSCcode = row.getCell(j).getStringCellValue();

								//validation started
								Date date = new Date();
								String formattedDate = dateFormat.format(date);
								Date date1 = dateFormat.parse(UTRDate);
								System.out.println("date1" + date1);
								Date date2 = dateFormat.parse(formattedDate);
								System.out.println("date2" + date2);
								int result = date1.compareTo(date2);
								System.out.println(result);
								if (result > 0) {
									// updaloade ddataalidation
									System.out.println("updaloade ddataalidation ERROR");
								} else {
									System.out.println("updaloade ddataalidation");
								}

							} else {
								// Failure case Data type is mismatch
								break;

							}

						}

					}

				}

			}

		}

	}
}

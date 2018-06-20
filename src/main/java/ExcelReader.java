import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader {

	public static final String excelFilePath = System.getProperty("user.dir") + "/src/main/resources/MyExcel.xlsx";

	public static void main(String[] args) throws IOException, InvalidFormatException {

		Workbook wb = WorkbookFactory.create(new File(excelFilePath));

		Sheet sheet = wb.getSheetAt(0);

		PrintWriter writer = new PrintWriter(System.getProperty("user.dir") + "/src/main/resources/ReadFromExcel.json",
				"UTF-8");

		writer.println("{");
		Iterator<Row> rowIterator = sheet.rowIterator();

		rowIterator.next();

		while (rowIterator.hasNext()) {
			writer.println("\t{");
			Row row = rowIterator.next();

			Iterator<Cell> cellIterator = row.cellIterator();

			int i = 0;
			while (cellIterator.hasNext()) {

				Cell cell = cellIterator.next();

				Row row1 = sheet.getRow(0);
				String coloumnName = row1.getCell(i).toString();

				switch (cell.getCellTypeEnum()) {
				case STRING:
					writer.print("\t\t" + "\"" + coloumnName + "\"" + ":" + "\"" + cell.getStringCellValue() + "\"");
					break;
				case NUMERIC:
					writer.print("\t\t" + "\"" + coloumnName + "\"" + ":" + cell.getNumericCellValue());
					break;
				case BOOLEAN:
					writer.print("\t\t" + "\"" + coloumnName + "\"" + ":" + cell.getBooleanCellValue());
					break;
				case FORMULA:
					writer.print("\t\t" + "\"" + coloumnName + "\"" + ":" + cell.getCellFormula());
					break;
				default:
					break;
				}
				i++;

				if (cellIterator.hasNext())
					writer.println(",");
			}

			writer.print("\n\t}");
			if (rowIterator.hasNext())
				writer.println(",");
		}

		writer.println("\n}");
		writer.close();
		System.out.println("Hi code executed successfully {}");
		
	}

}

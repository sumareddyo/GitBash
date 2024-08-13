package DDT;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Read_Excel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\issum\\OneDrive\\Desktop\\TestScriptData.xlsx");
		//Step 2:create work book i.e open the excel in read mode 
		 Workbook book = WorkbookFactory.create(fis);
		 
		//Step 3:get the sheet data into object
		 Sheet sheet = book.getSheet("Org");
		//Step 4:navigate to required row
		 Row row = sheet.getRow(1);
		//Step 5:navigate to required cell
		 Cell cell = row.getCell(2);
		 //Step 6:capture the data inside the cell
		 String excelData = cell.getStringCellValue();
		System.out.println(excelData);
		//step 7: close the work book
		book.close();
	}

}

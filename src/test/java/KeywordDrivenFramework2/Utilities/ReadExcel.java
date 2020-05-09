package KeywordDrivenFramework2.Utilities;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcel {
    XSSFWorkbook wbWorkbook;
    XSSFSheet shSheet;

    public void openSheet(String filePath) {
        FileInputStream fs;
        try {
            fs = new FileInputStream(filePath);
            wbWorkbook = new XSSFWorkbook(fs);
            shSheet = wbWorkbook.getSheet(Constants.sheetName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }  catch (IOException e) {
            e.printStackTrace();
        }
    }

    public String getValueFromCell(int iRowNumber, int iColNumber) {
        return shSheet.getRow(iRowNumber).getCell(iColNumber).getStringCellValue();
    }

    public int getRowCount() {
        return shSheet.getLastRowNum()+1;
    }

    public int getColumnCount() {
        return shSheet.getRow(0).getLastCellNum();
    }
}

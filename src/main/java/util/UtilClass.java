package util;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import java.io.*;

public class UtilClass
{
    public static void copyExelToWord() throws IOException {

        FileInputStream fiswb = new FileInputStream("input.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fiswb);

        XWPFDocument doc = new XWPFDocument();
        File file = new File("output.docx");
        file.createNewFile();

        XSSFSheet sheet = wb.getSheetAt(0);
        int LastRowNum = sheet.getLastRowNum();
        for(int i=0; i<=LastRowNum; i++)
        {
            XSSFRow currentRowSheet = sheet.getRow(i);
            int lastColumnNumber = sheet.getRow(0).getLastCellNum();
            XWPFTable table = doc.createTable(1,lastColumnNumber);
            XWPFTableRow currentRowDoc = table.getRow(0);

            for (int j = 0; j< lastColumnNumber;j++)
            {
                currentRowDoc.getCell(j).setWidth(Integer.toString(sheet.getColumnWidth(j)));
                String cellValue = currentRowSheet.getCell(j).toString();
                currentRowDoc.getCell(j).setText(cellValue);
            }
            doc.createParagraph().createRun().setText("\n");
        }

        FileOutputStream fopdoc = new FileOutputStream("output.docx");
        doc.write(fopdoc);
        System.out.println("Finished");
    }
}

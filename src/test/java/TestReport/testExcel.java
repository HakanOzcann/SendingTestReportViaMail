package TestReport;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class testExcel
{

    @Test(priority = 1)
    public void setCell()

    {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        Row firstRow = sheet.createRow(0);
        Row secondRow = sheet.createRow(1);

        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLUE.getIndex());
        style.setFont(font);

        CellStyle styleAlert = workbook.createCellStyle();
        styleAlert.setFillForegroundColor(IndexedColors.RED.getIndex());
        styleAlert.setFillPattern(CellStyle.SOLID_FOREGROUND);

        Row rowName = sheet.getRow(0);
        Cell cellName = rowName.getCell(1);

        Cell cellId = firstRow.createCell(0);
        cellId.setCellValue("12345");

        Cell cellEmail = firstRow.createCell(1);
        cellEmail.setCellValue("test@gmail.com");

        Cell cellAlertFirst = firstRow.createCell(4);

        Cell cellIdSecond = secondRow.createCell(0);
        cellIdSecond.setCellValue("12345");

        Cell cellEmailSecond = secondRow.createCell(1);
        cellEmailSecond.setCellValue("test@hotmail.com");

        Cell cellAlertSecond = secondRow.createCell(4);

        while (!cellId.getStringCellValue().isEmpty())
        {

            if (cellId.getStringCellValue().equals(cellEmail.getStringCellValue()))
            {
                cellId.setCellStyle(style);
                firstRow.removeCell(cellEmail);
                cellAlertFirst.setCellStyle(styleAlert);
            }

            if (cellId.getStringCellValue().equals(cellIdSecond.getStringCellValue()))
            {
                cellId.setCellStyle(style);
                secondRow.removeCell(cellIdSecond);
                cellAlertSecond.setCellStyle(styleAlert);
            }

            if (cellId.getStringCellValue().equals(cellEmailSecond.getStringCellValue()))
            {
                cellId.setCellStyle(style);
                secondRow.removeCell(cellEmailSecond);
                cellAlertSecond.setCellStyle(styleAlert);
            }

            break;

        }

        try (FileOutputStream fos = new FileOutputStream(new File("/Users/hakanozcan/Desktop/ApachePOI-Excel.xlsx")))
        {
            workbook.write(fos);
        }

        catch (IOException e)
        {
            e.printStackTrace();
        }

        System.out.println("Written successfully to the Excel File.");
    }

}
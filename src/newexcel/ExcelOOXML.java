/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package newexcel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

public class ExcelOOXML {
    
    public  void GenerarReporte(int Id) throws Exception{
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "Hoja excel");
        HSSFFont headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);

        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 1, 4);
        sheet.addMergedRegion(cellRangeAddress);
        HSSFRow dataRow1 = sheet.createRow(0);
        Cell cell0 = CellUtil.createCell(dataRow1, 0, "Nombre: ");
        Cell cell1 = CellUtil.createCell(dataRow1, 1, "Nombre Completa Elba");
        String[] headers = new String[]{
            "Fecha",
            "Entrada",
            "Salida",
            "Entrada",
            "Salida"
        };

        Object[][] data = new Object[][]{
            new Object[]{"2018-06-13", "15:42:40", "15:42:40", "15:42:40", "15:42:40"},
            new Object[]{"2018-06-13", "15:42:40", "15:42:40", "15:42:40", "15:42:40"},
            new Object[]{"2018-06-13", "15:42:40", "15:42:40", "15:42:40", "15:42:40"}};

        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        headerStyle.setFont(font);

        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());

        HSSFRow headerRow = sheet.createRow(1);
        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            HSSFCell cell = headerRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header);
        }
        int filaInicial = 2;

        for (int i = 0; i < data.length; ++i) {
            
            HSSFRow dataRow = sheet.createRow(filaInicial);
            filaInicial++;
            Object[] d = data[i];
            String product = (String) d[0];
            String price = (String) d[1];
            String link = (String) d[2];

            dataRow.createCell(0).setCellValue(product);
            dataRow.createCell(1).setCellValue(price);
            dataRow.createCell(2).setCellValue(link);
            dataRow.createCell(3).setCellValue((String) d[3]);
            dataRow.createCell(4).setCellValue((String) d[4]);
        }

        FileOutputStream file = new FileOutputStream("workbook.xls");
        workbook.write(file);
        file.close();
    }
}

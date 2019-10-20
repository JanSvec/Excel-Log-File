import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

class XlsxFile {
    private Workbook wb;
    private String name;

    /**
     * Instantiating XlsxFile class creates new excel (xlsx) file.
     * @param name name of the excel file, must end with .xlsx suffix.
     * @param columnNames array with the names of each column, will
     *                    take place at first row in a sheet.
     */
    XlsxFile(String name, String[] columnNames) throws IOException {
        this.name = name;
        wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row headerRow = sheet.createRow(0);
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        CellStyle headerCellStyle = wb.createCellStyle();
        headerCellStyle.setFont(headerFont);
        for (int i = 0; i < columnNames.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(columnNames[i]);
            headerCell.setCellStyle(headerCellStyle);
            sheet.autoSizeColumn(i);
        }
        FileOutputStream fileOut = new FileOutputStream(
                System.getProperty("user.dir") + "\\src\\main\\" + name);
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }

    /**
     * Method to be called to store new data in existing excel file.
     * Multiple calls of this method add data to new row in sheet.
     * @param reportData array with data to store.
     */
    void toXlsxFile(String[] reportData) throws IOException{
        FileInputStream fileIn = new FileInputStream(
                new File(System.getProperty("user.dir") + "\\src\\main\\" + name));
        wb = WorkbookFactory.create(fileIn);
        Sheet sheet = wb.getSheetAt(0);
        int rowCount = sheet.getLastRowNum();
        Row row = sheet.createRow(++rowCount);
        int cellIndex = 0;
        for (String data : reportData){
            row.createCell(cellIndex).setCellValue(data);
            cellIndex++;
        }
        for (int i = 0; i < reportData.length; i++){
            sheet.autoSizeColumn(i);
        }
        FileOutputStream fileOut = new FileOutputStream(
                System.getProperty("user.dir") + "\\src\\" + name);
        wb.write(fileOut);
        wb.close();
        fileOut.close();
    }
}

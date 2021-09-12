package sg.co.xia.writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sg.co.xia.dto.Employee;
import sg.co.xia.util.DateTimeUtil;
import sg.co.xia.util.FileUtil;

public class XlsWriter {

    @SneakyThrows
    public void generateExcel(final List<Employee> employeeList) {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("员工考勤");

        CellStyle workingInWeekendCellStyle = workbook.createCellStyle();
        workingInWeekendCellStyle.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
        workingInWeekendCellStyle.setFillPattern(CellStyle.ALIGN_LEFT);

        CellStyle abnormalCellStyle = workbook.createCellStyle();
        abnormalCellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
        abnormalCellStyle.setFillPattern(CellStyle.BIG_SPOTS);

        int rowNum = 0;
        for (Employee employee : employeeList) {
            Row row = sheet.createRow(rowNum++);

            int cellNum = 0;
            Cell acNoCell = row.createCell(cellNum++);
            acNoCell.setCellValue(employee.getAcNo());

            Cell nameCell = row.createCell(cellNum++);
            nameCell.setCellValue(employee.getName());

            Cell eariestTimeCell = row.createCell(cellNum++);
            eariestTimeCell.setCellValue(employee.getEarliestEnterTime().format(DateTimeUtil.DATE_TIME_FORMATTER));

            Cell lastLeaveTimeCell = row.createCell(cellNum++);
            lastLeaveTimeCell.setCellValue(employee.getLastLeaveTime().format(DateTimeUtil.DATE_TIME_FORMATTER));

        }

        autoSizeColumns(workbook);

        String fileFullPath = FileUtil.FILE_FOLDER + "Completed_" + FileUtil.FILE_NAME + FileUtil.FILE_EXTENSION;
        generateResultFile(workbook, fileFullPath);

        System.out.println("Proceed completed successful, file: " + fileFullPath);

    }

    private void autoSizeColumns(XSSFWorkbook workbook) {
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() > 0) {
                Row row = sheet.getRow(sheet.getFirstRowNum());
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex);
                }
            }
        }
    }

    private void generateResultFile(final XSSFWorkbook workbook, final String fileFullPath) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(new File(fileFullPath));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}

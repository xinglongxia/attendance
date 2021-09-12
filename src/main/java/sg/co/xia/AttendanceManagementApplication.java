package sg.co.xia;

import java.util.List;

import sg.co.xia.dto.Employee;
import sg.co.xia.reader.XlsReader;
import sg.co.xia.writer.XlsWriter;

public class AttendanceManagementApplication {

    public static void main(String[] args) {
        XlsReader xlsReader = new XlsReader();
        List<Employee> employeeList = xlsReader.readXlsFile();

        XlsWriter xlsWriter = new XlsWriter();
        xlsWriter.generateExcel(employeeList);
    }

}

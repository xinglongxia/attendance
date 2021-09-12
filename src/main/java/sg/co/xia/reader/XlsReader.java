package sg.co.xia.reader;

import java.io.File;
import java.io.FileInputStream;
import java.time.DayOfWeek;
import java.time.LocalDateTime;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.stream.Collectors;

import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import sg.co.xia.dto.Employee;
import sg.co.xia.util.KeyUtil;
import sg.co.xia.util.DateTimeUtil;
import sg.co.xia.util.FileUtil;

public class XlsReader {

    @SneakyThrows
    public List<Employee> readXlsFile() {
        FileInputStream file = new FileInputStream(new File(FileUtil.FILE_FOLDER + FileUtil.FILE_NAME + FileUtil.FILE_EXTENSION));

        //Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        //Get first/desired sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);

        //201_2020_10_20 -> employee
        Map<String, Employee> employeeMap = new TreeMap<>();

        sheet.forEach(row -> handleEachRow(row, employeeMap));

        file.close();

        return convertAndSortToList(employeeMap);
    }

    private void handleEachRow(final Row row, final Map<String, Employee> employeeMap) {
        //skip header
        if (row.getRowNum() == 0) {
            return;
        }

        String acNo = row.getCell(0).getStringCellValue();
        String employeeName = row.getCell(2).getStringCellValue();
        String clockInDateTimeString = row.getCell(3).getStringCellValue();

        LocalDateTime clockInDateTime = LocalDateTime.parse(clockInDateTimeString, DateTimeUtil.DATE_TIME_FORMATTER);

        //201-2020-11-26
        String key = KeyUtil.generateKey(acNo, clockInDateTime.getYear(), clockInDateTime.getMonth().getValue(), clockInDateTime.getDayOfMonth());

        if (employeeMap.containsKey(key)) {
            Employee employee = employeeMap.get(key);
            if (clockInDateTime.isBefore(employee.getEarliestEnterTime())) {
                employee.setEarliestEnterTime(clockInDateTime);
            } else if (clockInDateTime.isAfter(employee.getLastLeaveTime())) {
                employee.setLastLeaveTime(clockInDateTime);
            }
        } else {
            Employee employee = new Employee()
                    .setAcNo(acNo)
                    .setName(employeeName)
                    .setLocalDate(clockInDateTime.toLocalDate())
                    .setEarliestEnterTime(clockInDateTime)
                    .setLastLeaveTime(clockInDateTime)
                    .setAbsent(false)
                    .setWorkingInWeekend(clockInDateTime.getDayOfWeek() == DayOfWeek.SATURDAY || clockInDateTime.getDayOfWeek() == DayOfWeek.SUNDAY);

            employeeMap.put(key, employee);
        }
    }

    private List<Employee> convertAndSortToList(final Map<String, Employee> employeeMap) {
        Comparator<Employee> sortByAcNoComparator = Comparator.comparing(Employee::getAcNo);
        Comparator<Employee> sortByDateComparator = Comparator.comparing(Employee::getLocalDate);

        return employeeMap.values()
                .stream()
                .sorted(sortByAcNoComparator.thenComparing(sortByDateComparator))
                .collect(Collectors.toList());

    }
}

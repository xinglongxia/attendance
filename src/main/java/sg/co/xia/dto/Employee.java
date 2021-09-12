package sg.co.xia.dto;

import java.time.LocalDate;
import java.time.LocalDateTime;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain = true)
public class Employee {
    private String acNo;

    private String name;

    private LocalDate localDate;

    private LocalDateTime earliestEnterTime;

    private LocalDateTime lastLeaveTime;

    private boolean isAbsent;

    private boolean isWorkingInWeekend;

//    @Override
//    public String toString() {
//        return new StringJoiner(" ")
//            .add(acNo)
//            .add(name)
//            .add(earliestEnterTime.format(DateTimeUtil.DATE_TIME_FORMATTER))
//            .add(lastLeaveTime.format(DateTimeUtil.DATE_TIME_FORMATTER))
//            .toString();
//    }
}

package sg.co.xia.util;

import java.util.StringJoiner;

public class KeyUtil {

    public static String generateKey(final String acNo, final int year, final int month, final int day) {
        return new StringJoiner("_")
            .add(acNo).add(String.valueOf(year))
            .add(String.valueOf(month))
            .add(String.valueOf(day)).toString();
    }
}

package com.NAtools.util;

import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {

    public static Timestamp startOfDay(String date) throws Exception {
        String startOfDay = date + " 00:00:00";
        return convertToTimestamp(startOfDay);
    }

    public static Timestamp endOfDay(String date) throws Exception {
        String endOfDay = date + " 23:59:59";
        return convertToTimestamp(endOfDay);
    }

    private static Timestamp convertToTimestamp(String dateTime) throws Exception {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return new Timestamp(dateFormat.parse(dateTime).getTime());
    }

    public static String formatDateToISO(Date date, boolean startOfDay) {
        SimpleDateFormat isoFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

        if (startOfDay) {
            return isoFormat.format(date) + " 00:00:00";
        } else {
            return isoFormat.format(date) + " 23:59:59";
        }
    }
}

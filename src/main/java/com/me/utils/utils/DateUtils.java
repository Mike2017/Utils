package com.me.utils.utils;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class DateUtils {

    public static final String YYYYMMDDHHMMSS = "yyyyMMddHHmmss";

    public static String date2StrWithSpecifiedFormat(Date date, String pattern) {
        LocalDateTime localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
        return localDateTime.format(DateTimeFormatter.ofPattern(pattern));
    }

    public static String getMomentTime() {
        return date2StrWithSpecifiedFormat(new Date(), YYYYMMDDHHMMSS);
    }
}

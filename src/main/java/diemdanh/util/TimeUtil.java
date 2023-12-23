package diemdanh.util;

import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

public final class TimeUtil {

    public static Timestamp convertLocalDateTimeToTimestamp(LocalDateTime localDateTime) {
        // Chuyển đổi từ múi giờ Việt Nam sang múi giờ quốc tế (UTC)
        ZoneId viTimeZone = ZoneId.of("Asia/Ho_Chi_Minh");
        ZoneId utcTimeZone = ZoneId.of("UTC");

        Timestamp timestamp = Timestamp.from(localDateTime.atZone(viTimeZone).toInstant());
        timestamp = Timestamp.valueOf(timestamp.toLocalDateTime().atZone(utcTimeZone).toLocalDateTime());

        return timestamp;
    }

    public static LocalDateTime convertToVietnamTimeZone(Timestamp timestamp) {
        // Chuyển đổi từ Timestamp sang LocalDateTime
        LocalDateTime localDateTime = timestamp.toLocalDateTime();

        // Chuyển đổi múi giờ từ quốc tế sang Việt Nam
        ZoneId utcZone = ZoneId.of("UTC");
        ZoneId vietnamZone = ZoneId.of("Asia/Ho_Chi_Minh");

        return localDateTime.atZone(utcZone).withZoneSameInstant(vietnamZone).toLocalDateTime();
    }

    public static String formatLocalDateTime(LocalDateTime dateTime, String pattern) {
        // Định dạng LocalDateTime theo pattern
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
        return dateTime.format(formatter);
    }

}

package diemdanh.util;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import diemdanh.entity.CheckIn;
import diemdanh.entity.Student;

public final class WorkbookUtil {

    public static Workbook createDailyReportWorkbook(LocalDateTime now, List<Student> students, Statement statement) throws SQLException {
        Workbook workbook = new XSSFWorkbook();

        createAttendanceSheet(workbook, now, students, statement);
        createSwipeCardSheet(workbook, now, statement);

        return workbook;
    }

    private static void createAttendanceSheet(Workbook workbook, LocalDateTime now, List<Student> students, Statement statement) throws SQLException {
        Sheet sheet = workbook.createSheet("Điểm danh tháng " + TimeUtil.formatLocalDateTime(now, "MM-yyyy"));
        createHeaderRow(sheet, "STT", "Mã số KTX", "Họ và tên", "Phòng", "Khoá", "Trạng thái", "Chú thích");

        int rowIndex = 1;
        for (Student student : students) {
            CheckIn checkIn = DataAccess.getLastCheckinInDay(statement, student.getId(), now);
            Row row = sheet.createRow(rowIndex++);
            fillAttendanceDataRow(row, rowIndex - 1, student, checkIn);
        }
    }

    private static void createSwipeCardSheet(Workbook workbook, LocalDateTime now, Statement statement) throws SQLException {
        Sheet sheet = workbook.createSheet("Quẹt thẻ " + TimeUtil.formatLocalDateTime(now, "dd-MM-yyyy"));
        createHeaderRow(sheet, "STT", "Mã số KTX", "Họ và tên", "Phòng", "Khoá", "Trạng thái", "Chú thích");

        int rowIndex = 1;
        Map<Student, CheckIn> studentCheckInMap = DataAccess.getAllCheckInInDay(statement, now);
        for (Map.Entry<Student, CheckIn> entry : studentCheckInMap.entrySet()) {
            Student student = entry.getKey();
            CheckIn checkIn = entry.getValue();
            Row row = sheet.createRow(rowIndex++);
            fillSwipeCardDataRow(row, rowIndex - 1, student, checkIn);
        }
    }

    private static void createHeaderRow(Sheet sheet, String... headers) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
    }

    private static void fillAttendanceDataRow(Row row, int index, Student student, CheckIn checkIn) {
        int cellIndex = 0;
        row.createCell(cellIndex++).setCellValue(index);
        row.createCell(cellIndex++).setCellValue(student.getCode());
        row.createCell(cellIndex++).setCellValue(student.getFullname());
        row.createCell(cellIndex++).setCellValue(student.getRoom());
        row.createCell(cellIndex++).setCellValue("Khoá " + student.getGrade());
        if (checkIn.getType().equalsIgnoreCase("O")) {
            row.createCell(cellIndex++).setCellValue("Vắng mặt");
        } else {
            row.createCell(cellIndex++).setCellValue("Có mặt");
        }

        Cell cellNote = row.createCell(cellIndex++);
        Timestamp checkInTime = checkIn.getTime();
        String formattedTime = TimeUtil.formatLocalDateTime(TimeUtil.convertToVietnamTimeZone(checkInTime), "HH:mm");
        if (checkIn.getType().equalsIgnoreCase("O")) {
        		cellNote.setCellValue("Quẹt thẻ ra lần cuối vào lúc " + formattedTime);
        } else {
        		cellNote.setCellValue("Quẹt thẻ vào lần cuối vào lúc " + formattedTime);
        }
    }

    private static void fillSwipeCardDataRow(Row row, int index, Student student, CheckIn checkIn) {
        int cellIndex = 0;
        row.createCell(cellIndex++).setCellValue(index);
        row.createCell(cellIndex++).setCellValue(student.getCode());
        row.createCell(cellIndex++).setCellValue(student.getFullname());
        row.createCell(cellIndex++).setCellValue(student.getRoom());
        row.createCell(cellIndex++).setCellValue("Khoá " + student.getGrade());
        if (checkIn.getType().equalsIgnoreCase("O")) {
            row.createCell(cellIndex++).setCellValue("Vào");
        } else {
            row.createCell(cellIndex++).setCellValue("Ra");
        }
        
        Timestamp checkInTime = checkIn.getTime();
        String formattedTime = TimeUtil.formatLocalDateTime(TimeUtil.convertToVietnamTimeZone(checkInTime), "HH:mm");
        Cell cellMessage = row.createCell(cellIndex);
        if (checkIn.getType().equalsIgnoreCase("O")) {
            cellMessage.setCellValue("Quẹt thẻ ra lúc " + formattedTime);
        } else {
            cellMessage.setCellValue("Quẹt thẻ vào lúc " + formattedTime);
        }
    }

}

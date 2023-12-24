package diemdanh.util;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
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
    
    public static Workbook createMonthlyReportWorkbook(LocalDateTime now, List<Student> students, Statement statement) throws SQLException {
        Workbook workbook = new XSSFWorkbook();

        createAttendanceMonthlySheet(workbook, now, students, statement);

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

    private static void createAttendanceMonthlySheet(Workbook workbook, LocalDateTime now, List<Student> students, Statement statement) throws SQLException {
        Sheet sheet = workbook.createSheet("Điểm danh " + TimeUtil.formatLocalDateTime(now, "MM-yyyy"));
        String[] dateLabels = createDateLabels(21, 20);
        createHeaderRow(sheet, dateLabels);

        int rowIndex = 1;
        for (Student student : students) {
            Row row = sheet.createRow(rowIndex++);
            fillAttendanceMonthlyDataRow(statement, row, rowIndex - 1, student, 21, 20);
        }
    }

    private static String[] createDateLabels(int daysBefore, int daysAfter) {
        List<String> dateLabels = new ArrayList<>();
        
        dateLabels.add("STT");
        dateLabels.add("Mã số KTX");
        dateLabels.add("Họ và tên");
        dateLabels.add("Phòng");
        dateLabels.add("Khoá");

        LocalDateTime startDate = LocalDateTime.now().minusMonths(1).withDayOfMonth(daysBefore);
        LocalDateTime endDate = LocalDateTime.now().withDayOfMonth(daysAfter);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM");

        while (startDate.isBefore(endDate)) {
            dateLabels.add(startDate.format(formatter));
            startDate = startDate.plusDays(1);
        }
        
        dateLabels.add("Tổng số ngày có mặt");
        dateLabels.add("Tiền ăn nhận được");

        return dateLabels.toArray(new String[0]);
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

    private static void fillAttendanceMonthlyDataRow(Statement statement, Row row, int index, Student student, int daysBefore, int daysAfter) throws SQLException {
        int cellIndex = 0;
        row.createCell(cellIndex++).setCellValue(index);
        row.createCell(cellIndex++).setCellValue(student.getCode());
        row.createCell(cellIndex++).setCellValue(student.getFullname());
        row.createCell(cellIndex++).setCellValue(student.getRoom());
        row.createCell(cellIndex++).setCellValue("Khoá " + student.getGrade());
        


        LocalDateTime startDate = LocalDateTime.now().minusMonths(1).withDayOfMonth(daysBefore);
        LocalDateTime endDate = LocalDateTime.now().withDayOfMonth(daysAfter);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM");
        
        int numOfPresentDays = 0;

        while (startDate.isBefore(endDate)) {
            CheckIn checkIn = DataAccess.getLastCheckinInDay(statement, student.getId(), startDate);
  
            if (checkIn.getType().equalsIgnoreCase("O")) {
                row.createCell(cellIndex++).setCellValue("Vắng");
            } else {
            		numOfPresentDays++;
                Timestamp checkInTime = checkIn.getTime();
                String formattedTime = TimeUtil.formatLocalDateTime(TimeUtil.convertToVietnamTimeZone(checkInTime), "HH:mm");
                row.createCell(cellIndex++).setCellValue(formattedTime);
            }

            startDate = startDate.plusDays(1);
        }
        row.createCell(cellIndex++).setCellValue(numOfPresentDays);
        row.createCell(cellIndex++).setCellValue(numOfPresentDays * 40000);
        
    }

    private static void fillSwipeCardDataRow(Row row, int index, Student student, CheckIn checkIn) {
        int cellIndex = 0;
        row.createCell(cellIndex++).setCellValue(index);
        row.createCell(cellIndex++).setCellValue(student.getCode());
        row.createCell(cellIndex++).setCellValue(student.getFullname());
        row.createCell(cellIndex++).setCellValue(student.getRoom());
        row.createCell(cellIndex++).setCellValue("Khoá " + student.getGrade());
        if (checkIn.getType().equalsIgnoreCase("O")) {
            row.createCell(cellIndex++).setCellValue("Ra");
        } else {
            row.createCell(cellIndex++).setCellValue("Vào");
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

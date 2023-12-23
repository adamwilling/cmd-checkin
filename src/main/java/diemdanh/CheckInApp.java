package diemdanh;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.time.LocalDateTime;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import diemdanh.entity.Student;
import diemdanh.util.DataAccess;
import diemdanh.util.TimeUtil;
import diemdanh.util.WorkbookUtil;

public class CheckInApp {

    public static void main(String[] args) {
        String url = "jdbc:mysql://localhost:3306/diemdanh";
        String username = "root";
        String password = "1234567890";

        LocalDateTime now = LocalDateTime.now().minusDays(5);

        try (Connection connection = DriverManager.getConnection(url, username, password);
             Statement statement = connection.createStatement()) {

            List<Student> students = DataAccess.getAllStudent(statement);

            Workbook workbook = WorkbookUtil.createDailyReportWorkbook(now, students, statement);

            String formattedDateTime = TimeUtil.formatLocalDateTime(now, "dd-MM-yyyy");
            try (FileOutputStream fileOut = new FileOutputStream(formattedDateTime + ".xlsx")) {
                workbook.write(fileOut);
            }

            System.out.println("==============================Đã xong==============================");
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}

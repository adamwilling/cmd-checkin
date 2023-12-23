package diemdanh.util;

import java.sql.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

import diemdanh.entity.CheckIn;
import diemdanh.entity.Student;

public final class DataAccess {
	public static CheckIn getLastCheckinInDay(Statement statement, int studentId, LocalDateTime date) throws SQLException {
		StringBuilder sql = new StringBuilder();
		
		sql.append("SELECT MAX(created_at), type ");
		sql.append("FROM diemdanh.thienvu_comaydorm_checkins ");
		sql.append("WHERE student_id ='" + studentId + "' AND created_at BETWEEN '" + DateTimeFormatter.ofPattern("yyyy-MM-dd").format(TimeUtil.convertLocalDateTimeToTimestamp(date).toLocalDateTime().minusDays(1)) + "17:00:00' ");
		sql.append("AND '" + DateTimeFormatter.ofPattern("yyyy-MM-dd").format(TimeUtil.convertLocalDateTimeToTimestamp(date).toLocalDateTime()) + " 16:59:59' ");
		sql.append("GROUP BY type ORDER BY MAX(created_at) DESC LIMIT 1");
		ResultSet resultSet =  statement.executeQuery(sql.toString());
		
		CheckIn checkIn = new CheckIn();
		
		if (resultSet.next()) {
			checkIn.setType(resultSet.getString("type"));
			checkIn.setTime(resultSet.getTimestamp("MAX(created_at)"));
		} else {
			checkIn.setType("O");
			checkIn.setTime(null);
		}

		System.out.println(checkIn.getType());
		
		return checkIn;
	}
	
    public static Map<Student, CheckIn> getAllCheckInInDay(Statement statement, LocalDateTime date) throws SQLException {
		StringBuilder sql = new StringBuilder();

		sql.append("SELECT st.id, st.id_in_dorm, st.fullname, rm.name room, ck.created_at, ck.type ");
		sql.append("FROM diemdanh.thienvu_comaydorm_checkins ck ");
        sql.append("INNER JOIN diemdanh.thienvu_comaydorm_students st ON ck.student_id = st.id ");
        sql.append("INNER JOIN diemdanh.thienvu_comaydorm_rooms rm ON st.room_id = rm.id ");
		sql.append("WHERE ck.created_at BETWEEN '"
		        + DateTimeFormatter.ofPattern("yyyy-MM-dd").format(TimeUtil.convertLocalDateTimeToTimestamp(date).toLocalDateTime().minusDays(1)) + " 17:00:00' ");
		sql.append("AND '" + DateTimeFormatter.ofPattern("yyyy-MM-dd").format(TimeUtil.convertLocalDateTimeToTimestamp(date).toLocalDateTime()) + " 16:59:59' ");
        
		ResultSet resultSet = statement.executeQuery(sql.toString());

        List<Map.Entry<Student, CheckIn>> entryList = new ArrayList<>();

		while (resultSet.next()) {
			Student student = new Student();
			student.setId(resultSet.getInt("id"));
			student.setCode(resultSet.getString("id_in_dorm"));
			student.setFullname(resultSet.getString("fullname"));
			student.setRoom(resultSet.getString("room"));
			student.setGrade(Integer.parseInt(resultSet.getString("id_in_dorm").substring(0, 2)) - 16);
			
			CheckIn checkIn = new CheckIn();
			checkIn.setType(resultSet.getString("type"));
			checkIn.setTime(resultSet.getTimestamp("created_at"));
	        
            Map.Entry<Student, CheckIn> entry = Map.entry(student, checkIn);
            entryList.add(entry);
		}
		
        Map<Student, CheckIn> sortedStudentCheckInMap = new LinkedHashMap<>();
		
        for (Map.Entry<Student, CheckIn> entry : entryList) {
            sortedStudentCheckInMap.put(entry.getKey(), entry.getValue());
        }

        return sortedStudentCheckInMap;
    }
	
    public static List<Student> getAllStudent(Statement statement) throws SQLException {
		StringBuilder sql = new StringBuilder();
		
        sql.append("SELECT st.id, st.id_in_dorm, st.fullname, rm.name room ");
        sql.append("FROM diemdanh.thienvu_comaydorm_cards c ");
        sql.append("INNER JOIN diemdanh.thienvu_comaydorm_students st ON c.id = st.card_id ");
        sql.append("INNER JOIN diemdanh.thienvu_comaydorm_rooms as rm ON st.room_id = rm.id ");
        sql.append("WHERE rm.name <> '001' AND st.id_in_dorm NOT LIKE '201%' AND st.id_in_dorm NOT LIKE '202%' AND st.deleted_at IS NULL ");
        sql.append("ORDER BY st.id_in_dorm");
        
		ResultSet resultSet = statement.executeQuery(sql.toString());

		List<Student> students = new ArrayList<Student>();

		while (resultSet.next()) {
			Student student = new Student();
			student.setId(resultSet.getInt("id"));
			student.setCode(resultSet.getString("id_in_dorm"));
			student.setFullname(resultSet.getString("fullname"));
			student.setRoom(resultSet.getString("room"));
			student.setGrade(Integer.parseInt(resultSet.getString("id_in_dorm").substring(0, 2)) - 16);
			students.add(student);
		}
		
        // Sắp xếp danh sách theo phòng, sau đó theo mã số KTX
		students.sort(Comparator.comparing(Student::getGrade).thenComparing(Student::getRoom).thenComparing(Student::getCode));

        return students;
    }
}

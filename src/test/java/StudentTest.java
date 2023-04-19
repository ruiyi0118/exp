import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.Property;
import org.example.Student;

public class StudentTest {

    public static void main(String[] args) throws IOException {

        // 模拟生成学生集合
        List<Student> students = generateStudents();

        // 创建Excel文件
        String filePath = "D:/ProfessionalLearning/6thSemester/金山云训练营/知识导入/student.xlsx";
        File file = new File(filePath);
        if (!file.exists()) {
            file.createNewFile();
        }

        // 创建工作簿和工作表
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("学生信息");

        // 获取注解和字段列表
        List<Field> annotatedFields = new ArrayList<>();
        Field[] fields = Student.class.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(Property.class)) {
                annotatedFields.add(field);
            }
        }

        // 创建表头行
        Row headerRow = sheet.createRow(0);
        int cellIndex = 0;
        for (Field field : annotatedFields) {
            Property annotation = field.getAnnotation(Property.class);
            Cell cell = headerRow.createCell(cellIndex++);
            cell.setCellValue(annotation.name());
        }

        // 创建数据行
        int rowIndex = 1;
        for (Student student : students) {
            Row dataRow = sheet.createRow(rowIndex++);
            cellIndex = 0;
            for (Field field : annotatedFields) {
                try {
                    field.setAccessible(true);
                    Object value = field.get(student);
                    Cell cell = dataRow.createCell(cellIndex++);
                    if (value == null) {
                        cell.setCellType(CellType.BLANK);
                    } else if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Integer) {
                        cell.setCellValue((Integer) value);
                    } else if (value instanceof Date) {
                        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                        String dateValue = dateFormat.format((Date) value);
                        cell.setCellValue(dateFormat.parse(dateValue));
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }

        // 写入Excel文件并刷新输出流
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.flush();

        System.out.println("学生信息已输出到Excel文件：" + filePath);
    }

    /**
     * 模拟生成学生集合
     */
    private static List<Student> generateStudents() {
        List<Student> students = new ArrayList<>();
        students.add(new Student("张三", 18, "北京市海淀区", new Date()));
        students.add(new Student("李四", 20, "上海市浦东新区", new Date()));
        students.add(new Student("王五", 22, "广州市天河区", new Date()));
        return students;
    }

}

package org.example;

import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;

public class StudentTest {

    public static void main(String[] args) throws Exception {
        List<Student> students = generateStudents();

        String filePath = "D:\\ProfessionalLearning\\6thSemester\\金山云训练营\\知识导入\\students.xlsx";
        FileOutputStream fos = new FileOutputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("学生信息");

        // 输出表头
        XSSFRow headerRow = sheet.createRow(0);
        Field[] fields = Student.class.getDeclaredFields();
        int cellNum = 0;
        for (Field field : fields) {
            if (field.isAnnotationPresent(Property.class)) {
                Property property = field.getAnnotation(Property.class);
                XSSFCell cell = headerRow.createCell(cellNum++);
                cell.setCellValue(property.name());
            }
        }

        // 输出表格数据
        int rowNum = 1;
        for (Student student : students) {
            XSSFRow row = sheet.createRow(rowNum++);
            cellNum = 0;
            for (Field field : fields) {
                if (field.isAnnotationPresent(Property.class)) {
                    Object value = field.get(student);
                    XSSFCell cell = row.createCell(cellNum++);
                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Integer) {
                        cell.setCellValue((Integer) value);
                    } else if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                        XSSFCellStyle cellStyle = workbook.createCellStyle();
                        XSSFDataFormat dataFormat = workbook.createDataFormat();
                        cellStyle.setDataFormat(dataFormat.getFormat("yyyy-MM-dd"));
                        cell.setCellStyle(cellStyle);
                    } else {
                        cell.setCellType(CellType.BLANK);
                    }
                }
            }
        }

        workbook.write(fos);
        workbook.close();
        fos.close();
    }

    private static List<Student> generateStudents() {
        List<Student> students = new ArrayList<>();
        students.add(new Student("张三", 18, "北京市", new Date(2005 - 1900, 0, 1)));
        students.add(new Student("李四", 19, "上海市", new Date(2004 - 1900, 3, 1)));
        students.add(new Student("王五", 20, "广州市", new Date(2003 - 1900, 6, 1)));
        return students;
    }
}


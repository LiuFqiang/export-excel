import com.github.javafaker.Faker;
import com.github.javafaker.Name;
import dto.StudentDto;
import utils.ExportExcelUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class Application {

    public static void main(String[] args) throws IOException {
        List<StudentDto> studentDtoList = new ArrayList<>();
        for (int i = 0; i < 1000; i++) {
            StudentDto studentDto = new StudentDto();
            // 模拟数据
            Faker faker = new Faker();
            studentDto.setName(faker.name().name())
                    .setAddress(faker.address().fullAddress())
                    .setStudentId(faker.code().ean8());
            studentDtoList.add(studentDto);
        }

        InputStream is = null;
        try {
            is = ExportExcelUtil.export(studentDtoList, "sheet1", StudentDto.class);
            OutputStream outputStream = new FileOutputStream("C:\\Users\\15641\\Desktop\\111.xls");
            int ch = 0;
            while ((ch = is.read()) != -1) {
                outputStream.write(ch);
            }
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            is.close();
            System.out.println("Excel导出完成");
        }

    }
}

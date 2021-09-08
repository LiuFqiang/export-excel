package dto;

import annotation.Excel;

import java.io.Serializable;

public class StudentDto implements Serializable {

    @Excel(name = "学号")
    private String studentId;

    @Excel(name = "姓名", width = 30)
    private String name;

    @Excel(name = "家庭住址", width = 70)
    private String address;

    private String sex;

    public String getName() {
        return name;
    }

    public StudentDto setName(String name) {
        this.name = name;
        return this;
    }

    public String getStudentId() {
        return studentId;
    }

    public StudentDto setStudentId(String studentId) {
        this.studentId = studentId;
        return this;
    }

    public String getAddress() {
        return address;
    }

    public StudentDto setAddress(String address) {
        this.address = address;
        return this;
    }

    public String getSex() {
        return sex;
    }

    public StudentDto setSex(String sex) {
        this.sex = sex;
        return this;
    }
}

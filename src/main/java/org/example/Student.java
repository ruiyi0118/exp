package org.example;

import java.util.Date;

public class Student {
    @Property(name = "姓名")
    private String name;
    @Property(name = "年龄")
    private int age;
    @Property(name = "住址")
    private String address;
    @Property(name = "出生年月")
    private Date birthday;

    public Student(String name, int age, String address, Date birthday) {
        this.name = name;
        this.age = age;
        this.address = address;
        this.birthday = birthday;
    }
}


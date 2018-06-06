
package com.dianji.pangu.sdk.execl.model;

import com.alibaba.fastjson.JSONArray;
import com.dianji.pangu.sdk.excel.common.ErrorMessage;
import com.dianji.pangu.sdk.excel.common.ExcelDataFormatter;
import com.dianji.pangu.sdk.excel.common.FieldTree;
import com.dianji.pangu.sdk.excel.common.SheetConfig;
import com.dianji.pangu.sdk.excel.model.ExcelData;
import com.dianji.pangu.sdk.excel.util.ExcelReflectUtils;
import com.dianji.pangu.sdk.excel.util.ExcelUtils;
import com.dianji.pangu.sdk.excel.util.TypeReference;
import com.dianji.pangu.sdk.execl.entity.*;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

@SuppressWarnings("unused")
public class ExcelUtilsTest {

    //表示这个方法将提供数据给任何声明它的data provider名为“test1”的测试方法中
    @DataProvider(name = "test3")
    public Object[][] createData1() {
        Object[][] data = new Object[2][3];
        ExcelDataFormatter edf = new ExcelDataFormatter();
        Map<String, String> map = new HashMap<>();
        List<Student> list = makeUsersForTest();
        //需要双向
        map.put("true", "是");
        map.put("false", "否");
        map.put("是", "true");
        map.put("否", "false");
        edf.set("锁定Boolean", map);

        String destFile = getSystemPath() + "d1.xlsx";
        String srcFile = getSystemPath() + "d1.xlsx";
        data[0] = new Object[]{list, edf, srcFile, destFile};
        destFile = getSystemPath() + "d2.xlsx";
        srcFile = getSystemPath() + "d2.xlsx";
        data[1] = new Object[]{list, null, srcFile, destFile};

        return data;
    }

    /**
     * 制造测试数据 User对象列表
     *
     * @return java.model.List<com.dianji.pangu.sdk.execl.common.Person>
     */
    private List<Student> makeUsersForTest() {
        School school = new School();
        school.setSchoolName("深圳第一高级中学");
        school.setAddress("深圳市福田区芝麻公寓");
        SchoolMaster schoolMaster = new SchoolMaster("半仙", "男");
        school.setSchoolMaster(schoolMaster);
        List<Student> list = new ArrayList<>();

        Student u = new Student();
        u.setLocked(false);
        u.setTestFloat(2222f);
        u.setGrade(1);
        u.setSchool(school);
        list.add(u);

        u = new Student();
        u.setXx(123.23);
        u.setYy(getDateFromString("19980201000000"));
        u.setTime(getDateFromString("19980301120000"));
        u.setLocked(true);
        u.setDb(new BigDecimal("234.01"));
        u.setSchool(school);
        u.setScore(88.4);
        u.setClassNo(2);
        list.add(u);

        u = new Student();
        u.setAge(1234L);
        u.setName("用户1");
        u.setXx(123.23);
        u.setYy(getDateFromString("19880201000000"));
        u.setTime(getDateFromString("19880301120000"));
        u.setLocked(false);
        u.setDb(new BigDecimal(2344.00));
        u.setTestFloat(45646.0f);
        u.setScore(78.4);
        u.setSchool(school);
        list.add(u);

        u = new Student();
        u.setAge(22L);
        u.setName("用户二");
        u.setXx(123.23);
        u.setYy(getDateFromString("19830201000000"));
        u.setLocked(true);
        u.setDb(new BigDecimal(908));
        u.setTestFloat(45646.0f);
        u.setTime(getDateFromString("19780301120000"));
        u.setSchool(null);
        list.add(u);
        return list;
    }

    /**
     * 制造测试数据 User对象列表
     *
     * @return java.model.List<com.dianji.pangu.sdk.execl.common.Person>
     */
    private List<StudentT<Person, School>> makeStudentTsForTest() {
        School school = new School();
        school.setSchoolName("深圳第一高级中学");
        school.setAddress("深圳市福田区芝麻公寓");
        SchoolMaster schoolMaster = new SchoolMaster("半仙", "男");
        school.setSchoolMaster(schoolMaster);
        List<StudentT<Person, School>> list = new ArrayList<>();
        Person person = new Person();

        StudentT<Person, School> u = new StudentT<Person, School>();
        person.setLocked(false);
        person.setTestFloat(2222f);
        person.setXx(123.23);
        person.setYy(getDateFromString("19980201000000"));
        person.setTime(getDateFromString("19980301120000"));
        u.setName("name1");
        u.setX(10);
        u.setObj(person);
        u.setVobj(school);
        list.add(u);
        u = new StudentT<Person, School>();
        u.setName("name2");
        u.setX(12);
        u.setObj(person);
        u.setVobj(school);
        list.add(u);

        return list;
    }

    @Test
    public void testStudentTToFile() throws Exception {
        List<StudentT<Person, School>> list = makeStudentTsForTest();

        StringBuilder errormsg;
        TypeReference<StudentT<Person, School>> typeReference = new TypeReference<StudentT<Person, School>>() {
        };
        //测试工具封装类
        errormsg = ExcelUtils.writeToExcel(getSystemPath() + "/t1.xlsx", list, typeReference, null);

    }

    /**
     * 获取系统当前路径
     *
     * @return String 系统当前路径
     */
    private String getSystemPath() {
        Person u = new Person();
        return u.getClass().getResource("/").getFile();
    }

    /**
     * 根据字符串获取Data对象
     *
     * @param str 字符串Date，"yyyyMMddHHmmss"
     * @return Date对象
     */
    private Date getDateFromString(String str) {
        try {
            SimpleDateFormat s = new SimpleDateFormat("yyyyMMddHHmmss");
            return s.parse(str);
        } catch (Exception ex) {
            return null;
        }
    }

    /**
     * 比较两个list 指定行对象值是否完全相同
     *
     * @param srcPeople   源UserList
     * @param destPeople  目标UserList
     * @param srcStartNum 源开始行
     * @param dstStartNum 目标开始行
     * @param count       行数
     * @return Boolean
     */
    private Boolean compareUserList(List<Student> srcPeople, List<Student> destPeople, int srcStartNum, int dstStartNum, int count) {

        if (dstStartNum + count > destPeople.size()) {
            return false;
        }
        if (srcStartNum + count > srcPeople.size()) {
            return false;
        }
        for (int i = 0; i < count; i++) {
            System.out.println(srcPeople.get(i + srcStartNum).toString());
            System.out.println(destPeople.get(i + srcStartNum).toString());
//            if (!srcPeople.get(i+srcStartNum).toString().equals(destPeople.get(i+dstStartNum).toString())){
//                return false;
//            }
        }
        return true;
    }

    @Test(dataProvider = "test3")
    public void testWriteToFile(List<Student> list, ExcelDataFormatter edf, String srcFile, String destFile) throws Exception {
        SheetConfig sheetconfig = new SheetConfig();
        sheetconfig.setTitleRow(1);
        sheetconfig.setDataBeginRow(4);
        sheetconfig.setEdf(edf);
        Map<String, String> map = new HashMap<>();
        map.put("age", "年龄Long");
        map.put("xx", "Double测试");
        map.put("school", "");
        map.put("school.schoolName", "学校名称");
        map.put("school.address", "学校地址");
        map.put("school.schoolMaster", "");
        map.put("school.schoolMaster.name", "校长姓名");
        map.put("school.schoolMaster.sex", "校长性别");
        map.put("name", "姓名String");
//        sheetconfig.setFieldNameMap(map);
        OutputStream outputStream = new FileOutputStream(destFile);
        StringBuilder errormsg;
        //测试工具封装类
//      errormsg = ExcelUtils.writeToExcel(destFile,list,sheetconfig);
        //测试多次写
        ExcelData writeExcel = ExcelData.open(outputStream);
        try {
            errormsg = writeExcel.write(list, null, sheetconfig);
            System.out.print(errormsg.toString());
            errormsg = writeExcel.write(list, null, sheetconfig);
            System.out.print(errormsg.toString());
        } finally {
            writeExcel.close();
        }
//        Assert.assertEquals(DigestUtils.md5Hex(new FileInputStream(srcFile)), (DigestUtils.md5Hex(new FileInputStream(destFile))));

    }

    @Test(dataProvider = "test3")
    public void testReadExcel(List<Student> srcList, ExcelDataFormatter edf, String srcFileName, String destFileName) throws Exception {

        List<ErrorMessage> errorMsgs = new ArrayList<>();
        SheetConfig sheetconfig = new SheetConfig();
        sheetconfig.setTitleRow(1);
        sheetconfig.setDataBeginRow(4);
        sheetconfig.setEdf(edf);
        Map<String, String> map = new HashMap<>();
        map.put("age", "年龄Long");
        map.put("xx", "Double测试");
        map.put("school", "");
        map.put("school.schoolName", "学校名称");
        map.put("school.address", "学校地址");
        map.put("school.schoolMaster", "");
        map.put("school.schoolMaster.name", "校长姓名");
        map.put("school.schoolMaster.sex", "校长性别");
        map.put("name", "姓名String");
//        sheetconfig.setFieldNameMap(map);
        List<Student> destList;
        destList = ExcelUtils.readFromExcel(srcFileName, Student.class, sheetconfig, errorMsgs);
        for (ErrorMessage errorMessage : errorMsgs) {
            System.out.println("行：" + errorMessage.getRow() + errorMessage.getMsg().toString());
        }
        for (Student student : destList) {
            System.out.println(student.toString());

        }
        JSONArray errorJson = new JSONArray();
        destList = ExcelUtils.readFromExcel(srcFileName, Student.class, sheetconfig, errorJson);
        System.out.println(errorJson.toString());
        List<StudentT<Person, School>> studentTS;
        TypeReference<StudentT<Person, School>> typeReference = new TypeReference<StudentT<Person, School>>() {
        };
        studentTS = ExcelUtils.readFromExcel(srcFileName, typeReference, sheetconfig, errorMsgs);
        for (StudentT studentT : studentTS) {
            if (studentT.getVobj() != null) {
                System.out.println(studentT.getVobj().toString());
            }
            if (studentT.getObj() != null) {
                System.out.println(studentT.getObj().toString());
            }

        }


//       Assert.assertTrue(compareUserList(srcList,destList,0,0,4));
    }

    @Test
    public void testReadExcelToJson() throws Exception {
        JSONArray data;
        String srcFileName = getSystemPath() + "/s2.xlsx";
        data = ExcelUtils.readFromExcel(srcFileName, null);
        System.out.println(data.toString());
    }

    @Test
    public void testFieldsType() throws Exception {

        StudentT<Student, Person> s = new StudentT<>();
        Field[] fields;
        fields = ExcelReflectUtils.getClassAllDeclaredFields(StudentT.class);
        for (Field field : fields) {
            if (field.getType().equals(Integer.class)) {
                System.out.println(field.getName() + " Class is :" + field.getType().getName()
                        + ", Declaring Class: " + field.getDeclaringClass().getName()
                        + ", GetGenericType: " + field.getGenericType().getTypeName());
            }
            if (field.getType().isPrimitive()) {
                System.out.println(field.getName() + " Class is :" + field.getType().getName()
                        + ", Declaring Class: " + field.getDeclaringClass().getName()
                        + ", GetGenericType: " + field.getGenericType().getTypeName());
            }
            System.out.println(field.getName() + " Class is :" + field.getType().getName()
                    + ", Declaring Class: " + field.getDeclaringClass().getName()
                    + ", GetGenericType: " + field.getGenericType().getTypeName());
            field.setAccessible(true);
            Type t = field.getGenericType();

            if (t instanceof ParameterizedType) {
                Class c = (Class<?>) ((ParameterizedType) t).getActualTypeArguments()[0];
                System.out.println(c.getName());
            }
            Object value = field.get(s);
            if (value instanceof Integer) {
                System.out.println(field.getName() + " Class is :" + field.getType().getName()
                        + ", Declaring Class: " + field.getDeclaringClass().getName()
                        + ", GetGenericType: " + field.getGenericType().getTypeName());

            }
        }
    }

    @Test
    @SuppressWarnings("unchecked")
    public void testFieldsTree() throws Exception {
        TypeReference<StudentT<Student, Person>> typeReference = new TypeReference<StudentT<Student, Person>>() {
        };
//        TypeReference<Student> typeReference = new TypeReference<Student>(){};
        FieldTree root;
        root = ExcelReflectUtils.makeFieldTree(typeReference, null);
        StringBuilder sb = new StringBuilder();
        ExcelReflectUtils.printTree(root, sb);
        StudentT<Student, Person> st = (StudentT<Student, Person>) root.getFieldType().newInstance();
        StudentT<Student, Person> st1 = new StudentT<>();
        System.out.println(sb.toString());
        System.out.println("--------------------------------------------------------");
/*
        Student s=new Student();
        root = ExcelReflectUtils.makeFieldTree(Student.class,null);
        sb = new StringBuilder();
        ExcelReflectUtils.printTree(root,sb);
        System.out.println(sb.toString());
        System.out.println("--------------------------------------------------------");
        String[] titleNames={"姓名String","学校名称","学校地址","年级","班级","成绩",};
        List<Object> list = new ArrayList<>();
        list.add("hanson");
        list.add("深圳大学");
        list.add("深圳南山区");
        list.add(2);
        list.add(3L);
        list.add(78.89);
        root = ExcelReflectUtils.makeFieldTree(Student.class,titleNames,null);
        sb.setLength(0);
        ExcelReflectUtils.printTree(root,sb);
        System.out.println(sb.toString());
        ExcelReflectUtils.listToBean(list,root,s);
        System.out.println("姓名:"+s.getName());
        System.out.println("学校名称:"+s.getSchool().getSchoolName());
        System.out.println("学校地址:"+s.getSchool().getAddress());
        System.out.println("年级:"+s.getGrade());
        System.out.println("班级:"+s.getClassNo());
        System.out.println("成绩:"+s.getScore());
        */
    }

} 

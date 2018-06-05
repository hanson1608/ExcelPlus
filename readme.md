ExcelPlus
===
ExcelPlus is a java library that can be used to convert the data in excel file to Jave Objects.You can read the data from excel file to Jave Objects and also can write the Java Object to Excel file.

* intergrate the java valification <br> 
* support the field of Java Object which is user define class <br> 
* support both the annotation and config  <br> 
* one sheet to one Java Object <br> 
* support generic type <br> 

How to use generic type
---

        //泛型请调用TypeReference参数接口。务必：new TypeReference<StudentT<Person,School>>(){}方式，生成匿名子类
        //写示例
        TypeReference<StudentT<Person,School>> typeReference = new TypeReference<StudentT<Person,School>>(){};
        //测试工具封装类
        errormsg = ExcelUtils.writeToExcel(getSystemPath()+"/t1.xlsx",list,typeReference,null);
        //读
        List<StudentT<Person,School>> studentTS;
        TypeReference<StudentT<Person,School>> typeReference = new TypeReference<StudentT<Person,School>>(){};
        studentTS =ExcelUtils.readFromExcel(srcFileName,typeReference,sheetconfig,errorMsgs);




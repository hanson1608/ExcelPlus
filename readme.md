修改历史:
2018年5月22日提交
对读excel文件到实体对象修改调整
1、增加配置类型：明确数据区和标题区 

2、数据校验功能独立增强：集成支持java valification。注解接口模块实现对数据读取过程中的更强大灵活校验，校验功能从excel注解里移除 

3、通过构建字段映射树增加对实体类中子类的嵌套支持。 

4、增加可以在一个excel上读写多个sheet，配置基于sheet配置，每个sheet数据映射到不同的实体对象 

5、重新调整代码组织、结构和接口 

    5.1、结构上拆分ExcelData、SheetData对象,ExcelUtils类完成常用访问的包装，灵活的方式可以直接对ExcelData和SheetData进行包装。
    5.2、底层把数据转换统一拆分成open，多次read/write,close的统一操作路径，根据open的时候时输入流还是输出流自动判断时都还是写。
    5.3、扩展了反射类ReflectUtils为ExcelReflectUtils，把字段映射树FieldTree的构建在里面完成。
    5.4、根据功能重新把目录结构调整为common(通用如注解、配置类、字段映射树类）,model（对excel读写映射的业务实现）,util（内外的工具类接口，原则上都是静态方法）
    
6、数据类型支持扩展：同时支持了基本类型和他们的包装类 

7、增加对类标志名和sheet名的约定匹配。 

    7.1、读的时候：根据类的注解自动匹配对应的sheet。约定匹配规则：先按照类的注解找到对应的sheet，找不到缺省匹配为第一个sheet
    7.2、写的时候：当多个实例类写到不同sheet里，sheet名称重复时自动后面加当前Sheet总数进行区分
    
8、修正了对涉及到的相关资源的关闭 

9、完善了读写过程中错误信息的收集反馈到上层调用者。 

10、所有的类必须有无参的构造方法 

11、修正了bug

    11.1、行中间单元格跳跃的情况出错的bug，excel只保存有内容（包括空）的单元格
    11.2、浮点数字转换错误
    
12、注意不支持char和byte类型。

###2018年5月24日
1、支持从第三方把配置传过来的方式，替代注解约定.支持类下面的字段名（字段名.字段名.XXX)方式。字段父字段也必须有对应的配置，否则向下都忽略。

###2018年5月28日
1、修正了类没有注解出错问题、增加了对stream的关闭。
2、增加了对泛型的支持

        //泛型请调用TypeReference参数接口。务必：new TypeReference<StudentT<Person,School>>(){}方式，生成匿名子类
        //写示例
        TypeReference<StudentT<Person,School>> typeReference = new TypeReference<StudentT<Person,School>>(){};
        //测试工具封装类
        errormsg = ExcelUtils.writeToExcel(getSystemPath()+"/t1.xlsx",list,typeReference,null);
        //读
        List<StudentT<Person,School>> studentTS;
        TypeReference<StudentT<Person,School>> typeReference = new TypeReference<StudentT<Person,School>>(){};
        studentTS =ExcelUtils.readFromExcel(srcFileName,typeReference,sheetconfig,errorMsgs);




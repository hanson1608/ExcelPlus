/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: ReflectUtils
 * Author:   hanson
 * Date:     2018/5/19 15:05
 *
 * @since 1.0.0
 */
package com.dianji.pangu.sdk.excel.util;

import com.dianji.pangu.sdk.excel.common.ErrorMessage;
import com.dianji.pangu.sdk.excel.common.Excel;
import com.dianji.pangu.sdk.excel.common.ExcelException;
import com.dianji.pangu.sdk.excel.common.FieldTree;

import java.lang.reflect.*;
import java.math.BigDecimal;
import java.util.*;

/**
 * ExcelReflectUtils Excel反射工具类,包括通过反射构建字段映射树
 * @author hanson
 * @date 2018/5/19
 * @since 1.0.0
 */
@SuppressWarnings("unused")
public class ExcelReflectUtils {

    private ExcelReflectUtils(){
    }
    /**
     * 获取成员变量(包括private,protected)的修饰符
     * @param clazz 类
     * @param fieldName 字段名
     * @return int
     */
    public static <T> int getFieldModifier(Class<T> clazz, String fieldName)  {
        //getDeclaredFields可以获取所有修饰符的成员变量，包括private,protected等getFields则不可以
        Field[] fields = clazz.getDeclaredFields();

        for (Field field:fields) {
            if (field.getName().equals(fieldName)) {
                return field.getModifiers();
            }
        }
        return 0;
    }
    /**
     * 获取成员方法的修饰符
     * @param clazz 类
     * @param methodName 方法名
     * @return int
     */
    public static <T> int getMethodModifier(Class<T> clazz, String methodName)  {

        //getDeclaredMethods可以获取所有修饰符的成员方法，包括private,protected等getMethods则不可以
        Method[] methods = clazz.getDeclaredMethods();

        for (Method method:methods) {
            if (method.getName().equals(methodName)) {
                return method.getModifiers();
            }
        }
        return 0;
    }

    /**
     * [对象]根据成员变量名称获取其值
     * @param entity 类实例
     * @param fieldName 字段名
     * @return Object
     */
    public static Object getFieldValue(Object entity, String fieldName) throws   IllegalAccessException {
        Field[] fields = entity.getClass().getDeclaredFields();

        for (Field field:fields) {
            if (field.getName().equals(fieldName)) {
                //对于私有变量的访问权限，在这里设置，这样即可访问Private修饰的变量
                field.setAccessible(true);
                return field.get(entity);
            }
        }
        return null;
    }
    /**
     * [类]根据成员变量名称获取其字段缺省默认值
     * @param clazz 类
     * @param fieldName 字段名
     * @return Object
     */
    public static <T> Object getFieldValue(Class<T> clazz, String fieldName) throws   IllegalAccessException, InstantiationException {
        T clazzInstance = clazz.newInstance();
        return getFieldValue(clazzInstance,fieldName);
    }
    /**
     * 获取该类下声明所有的成员变量名称（包含private，public，protected）
     * @param clazz 类
     * @return java.lang.String[]
     */
    public static  String[] getFields(Class clazz) {
        Field[] fields = clazz.getDeclaredFields();
        String[] fieldsArray = new String[fields.length];

        for (int i = 0; i < fields.length; i++) {
            fieldsArray[i] = fields[i].getName();
        }
        return fieldsArray;
    }
    /**
     * 获取这个类所有包括（父类，祖先，到object类为止）所有申明的字段（public，protected，private）
     * @param clazz 对象类
     * @return java.lang.reflect.Field[]
     */
    public static  Field[] getClassAllDeclaredFields(Class clazz) {
        List<Field> allFieldList = new ArrayList<>();

        for (Class<?> curClass = clazz; curClass!=Object.class; curClass= curClass.getSuperclass()){
            Field[] fields = curClass.getDeclaredFields();
            allFieldList.addAll(Arrays.asList(fields));
        }

        return allFieldList.toArray(new Field[allFieldList.size()]);
    }

    /**
     * 指定类，调用缺省的无参方法
     * @param clazz  类
     * @param methodName  方法名
     * @return java.lang.Object
     */
    public static <T> Object invoke(Class<T> clazz, String methodName)
            throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
        Object instance = clazz.newInstance();
        Method m = clazz.getMethod(methodName);
        return m.invoke(instance);
    }

    /**
     * 通过对象，访问其方法,无参数
     * @param clazzInstance 类实例对象
     * @param methodName 方法名
     */
    public static  Object invoke(Object clazzInstance, String methodName)
            throws NoSuchMethodException,  IllegalAccessException,  InvocationTargetException {
        Method m = clazzInstance.getClass().getMethod(methodName);
        return m.invoke(clazzInstance);
    }

    /**
     * 指定类，调用指定的方法
     * @param clazz  类
     * @param method  方法名
     * @param paramClasses 方法对应的参数类数组
     * @param params  方法对应的参数数组
     * @return java.lang.Object
     */
    public static <T> Object invoke(Class<T> clazz, String method, Class<T>[] paramClasses, Object[] params)
            throws InstantiationException, IllegalAccessException, NoSuchMethodException,  InvocationTargetException {
        Object instance = clazz.newInstance();
        Method m = clazz.getMethod(method, paramClasses);
        return m.invoke(instance, params);
    }

    /**
     * 通过类的实例，调用指定的方法
     * @param clazzInstance  类的实例
     * @param methodName  方法名
     * @param paramClasses 方法对应的参数类数组
     * @param params  方法对应的参数数组
     * @return java.lang.Object
     */
    public static <T> Object invoke(Object clazzInstance, String methodName, Class<T>[] paramClasses, Object[] params)
            throws  IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        Method method = clazzInstance.getClass().getMethod(methodName, paramClasses);
        return method.invoke(clazzInstance, params);
    }

    //和Excel fieldTree相关的方法
    /**
     * 根据excel注解构建字段和excel sheet中列名的对应关系
     * @param entityClass 实体类
     * @param titleNames 列名
     * @return FieldTree 构建完的完整的树
     */
    public static FieldTree makeFieldTree(Class entityClass,String[] titleNames,Map<String,String> fieldMap)throws ExcelException{
        FieldTree root = makeFieldTree(entityClass,fieldMap);
        adjustFieldTree(root,titleNames);
        return root;
    }
    /**
     * 根据excel注解构建字段和excel sheet中列名的对应关系
     * @param entityClass 实体类
     * @return FieldTree 构建完的完整的树
     */
    public static FieldTree makeFieldTree(Class entityClass,Map<String,String> fieldMap) throws ExcelException{
        //初始化建立root节点。root节点标志是field为null
        FieldTree root = new FieldTree();
        root.setField(null);
        root.setColTitleName(null);
        root.setBaseClassType(false);
        root.setFieldFullName("");
        root.setFieldType(entityClass);
        //根据field构建所有的字节点
        makeFieldTree(entityClass,root,0,fieldMap,null);
        return root;
    }
    /**
     * 根据excel注解构建字段和excel sheet中列名的对应关系(针对泛型)
     * @param typeReference 实体类引用包装，解决泛型的问题
     * @param titleNames 列名
     * @return FieldTree 构建完的完整的树
     */
    public static FieldTree makeFieldTree(TypeReference typeReference,String[] titleNames,Map<String,String> fieldMap)throws ExcelException{
        FieldTree root = makeFieldTree(typeReference,fieldMap);
        adjustFieldTree(root,titleNames);
        return root;
    }
    /**
     * 根据excel注解构建字段和excel sheet中列名的对应关系
     * @param typeReference 实体类引用包装，解决泛型的问题
     * @return FieldTree 构建完的完整的树
     */
    @SuppressWarnings("unchecked")
    public static FieldTree makeFieldTree(TypeReference typeReference,Map<String,String> fieldMap) throws ExcelException{
        //初始化建立root节点。root节点标志是field为null
        FieldTree root = new FieldTree();
        root.setField(null);
        root.setColTitleName(null);
        root.setBaseClassType(false);
        root.setFieldFullName("");
        root.setFieldType(typeReference.getRawType());
        Map<String,Class<?>> genericMap = typeReference.makeGenericMap();
        //根据field构建所有的字节点
        makeFieldTree(root.getFieldType(),root,0,fieldMap,genericMap);
        return root;
    }
    /**
     * 根据字段映射表构建字段和excel sheet中列名的对应关系
     * @param entityClass 实体类
     * @param parentFieldTreeNode 父节点
     * @param startCount 计数起始值
     * @param fieldMap 字段映射表
     * @return 计数
     */
    @SuppressWarnings("unchecked")
    private static int makeFieldTree(Class entityClass, FieldTree parentFieldTreeNode,int startCount,Map<String,String> fieldMap,Map<String,Class<?>> genericTypeMap) throws ExcelException{
        Field[] entityFields = getClassAllDeclaredFields(entityClass);
        //判断是否有缺省的构造方法，防止在构造的时候失败
        try{
            if (!parentFieldTreeNode.isBaseClassType()) {
                entityClass.getConstructor();
            }}catch (NoSuchMethodException ex){
            throw new ExcelException(entityClass+ ErrorMessage.NO_CONSTRUCTOR_WITHOUT_PARA);
        }
        //缺省从注解中定义对应关系
        int count = startCount;
        String parentFieldName = parentFieldTreeNode.getFieldFullName();
        for (Field field : entityFields) {
            //设置字段可以访问，便于直接操作私有字段，在这边统一设，不要在写的循环里设置
            field.setAccessible(true);
            //根据注解或字段Map构建字段映射树节点
            FieldTree childFieldTreeNode = makeFieldTreeNode(field,parentFieldName,fieldMap);
            if (childFieldTreeNode == null){
                continue;
            }
            //基本类型或包装类就直接使用，否则要深入该类看该类的字段是否有注解和Excel关联
            childFieldTreeNode.setBaseClassType(isExcelBaseType(field.getType()));
            //判断父节点是否是根，决定是否有前缀
            if(parentFieldName.isEmpty()){
                childFieldTreeNode.setFieldFullName(field.getName());
            }else {
                childFieldTreeNode.setFieldFullName(parentFieldName + "." + field.getName());
            }
            if(!childFieldTreeNode.isBaseClassType()){
                //设置为对应的列Index为-1，表示该字段不对应excel sheet的列
                childFieldTreeNode.setColIndex(-1);
                //支持对泛型的转换到实际类
                childFieldTreeNode.setFieldType(getFieldTrueType(field,genericTypeMap));
                count = makeFieldTree(childFieldTreeNode.getFieldType(),childFieldTreeNode,count,fieldMap,genericTypeMap);
            }else{
                childFieldTreeNode.setFieldType(field.getType());
                childFieldTreeNode.setColIndex(count);
                count++;
            }
            //加入到父亲节点中
            parentFieldTreeNode.addChild(childFieldTreeNode);
        }
        return count;
    }

    /**
     * 根据field和泛型Map找到field真实的类
     * @param field 字段
     * @param genericTypeMap 泛型类map
     * @return field代表的真正的类
     */
    private static Class<?> getFieldTrueType(Field field,Map<String,Class<?>> genericTypeMap ){
        if(genericTypeMap==null){
            return field.getType();
        }
        Type type = field.getGenericType();
        if (type instanceof Class<?>){
            return (Class<?>)type;
        }
        //泛型变量
        else if(type instanceof  TypeVariable){
            return genericTypeMap.get(type.getTypeName());
        }
        return Object.class;
    }

    /**
     * 根据注解或字段映射构建字段映射树
     * @param field 字段
     * @param parenetFieldName 父字段的全名
     * @param fieldMap 字段映射map
     * @return 字段映射树节点
     */
    private static FieldTree makeFieldTreeNode(Field field,String parenetFieldName,Map<String,String> fieldMap){
        String title;
        int width=20;
        String format="";

        FieldTree fieldTreeNode = new FieldTree();
        //外部配置map优先于注解，有外部配置就不考虑注解
        if(fieldMap == null) {
            Excel excelAnnotation = field.getAnnotation(Excel.class);
            if (excelAnnotation == null) {
                return null;
            }
            title = excelAnnotation.title();
            width = excelAnnotation.width();
            format = excelAnnotation.format();
        }else{
            String fullName=field.getName();
            if(!parenetFieldName.isEmpty()){
                fullName = parenetFieldName+'.'+field.getName();
            }
            title = fieldMap.get(fullName);
            if(title == null){
                return null;
            }
        }
        fieldTreeNode.setField(field);
        fieldTreeNode.setColTitleName(title);
        fieldTreeNode.setColWidth(width);
        fieldTreeNode.setColFormat(format);
        return fieldTreeNode;
    }

    /**
     * 根据excel sheet的列标题设置数和列号对应关系，索引小于0的叶子节点以后不管
     * @param fieldTree 字段对应树
     * @param titleNames 列名
     */
    private static void adjustFieldTree(FieldTree fieldTree,String[] titleNames){
        //叶子节点
        if (fieldTree.isBaseClassType()){
            fieldTree.setColIndex(getIndex(fieldTree.getColTitleName(),titleNames));
        }else{
            //遍历树
            List<FieldTree> children = fieldTree.getChildren();
            for (FieldTree child:children) {
                adjustFieldTree(child,titleNames);
            }
        }
    }

    /**
     * 根据str的值在str数组里找到对应的位置
     * @param str  需要找的字符串
     * @param strArray 字符串数组
     * @return 字符串所在的索引位置，<0表示没找到
     */
    private static int getIndex(String str,String[] strArray){
        for(int index = 0;index<strArray.length;index++){
            if (strArray[index].equals(str)){
                return index;
            }
        }
        return -1;
    }
    /**
     * 判断该类是不是包装类
     * @param clz 类
     * @return boolean 是否是包装类
     */
    @SuppressWarnings("all")
    public static boolean isWrapClass(Class clz) {
        try {
            return ((Class) clz.getField("TYPE").get(null)).isPrimitive();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * 判断该类是否是包装类或基本类,String,Date,BigDecimal
     * @param entityClass 类
     * @return boolean 是否是包装类或基本类
     */
    private static boolean isExcelBaseType(Class entityClass) {
        return (entityClass.isPrimitive()|| isWrapClass(entityClass)
                || entityClass.equals(BigDecimal.class) || entityClass.equals(String.class)
                || entityClass.equals(Date.class));
    }

    /**
     * 将list里的数据根据映射数构建实体对象
     * @param list 数据列表
     * @param fieldTree 映射树
     * @param currentObj 对象
     */
    public static void listToBean(List list, FieldTree fieldTree, Object currentObj) throws Exception{
        Field field;
        if (fieldTree.getField() != null) {
            field = fieldTree.getField();

            if (fieldTree.isBaseClassType()) {
                //基本类型就直接赋值
                int index = fieldTree.getColIndex();
                if (index>=0) {
                    field.set(currentObj,list.get(index));
                }
            }else{ //如果不是基本类型就往下构建
                //创建字段对象
                Object fieldObj = field.getType().newInstance();
                //往下构建字段对象
                for (FieldTree child : fieldTree.getChildren()) {
                    listToBean(list,child,fieldObj);
                }
                field.set(currentObj,fieldObj);
            }
        }else{//root
            //遍历子节点
            for (FieldTree child : fieldTree.getChildren()) {
                listToBean(list,child,currentObj);
            }
        }
    }

    /**
     * 打印生成的字段映射树，测试用
     * @param fieldTree 字段映射树
     */
    public static void printTree(FieldTree fieldTree,StringBuilder sb){
        //非根节点内容
        if(fieldTree.getField()!=null){
            sb.append("Name:");
            sb.append(fieldTree.getColTitleName());
            sb.append(",Width:");
            sb.append(fieldTree.getColWidth());
            sb.append(",Format:");
            sb.append(fieldTree.getColFormat());
            sb.append(", Index:");
            sb.append(fieldTree.getColIndex());
            sb.append(", Type:");
            sb.append(fieldTree.getField().getType());
            sb.append(", BaseType:");
            sb.append(fieldTree.isBaseClassType());
            sb.append(", FullName:");
            sb.append(fieldTree.getFieldFullName());
            sb.append(", ActualType:");
            sb.append(fieldTree.getFieldType().getName());
            sb.append("\n");
        }else{
            sb.append("root\n");
        }
        if (!fieldTree.isBaseClassType()){
            //遍历所有子节点
            for (FieldTree child:fieldTree.getChildren()) {
                printTree(child,sb);
            }
        }

    }
}

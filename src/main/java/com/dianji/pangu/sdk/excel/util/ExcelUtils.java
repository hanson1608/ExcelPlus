/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: ExcelUtils
 * Author:   hanson
 * Date:   2018/5/21 21:48
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.util;

import com.dianji.pangu.sdk.excel.common.ErrorMessage;
import com.dianji.pangu.sdk.excel.common.ExcelException;
import com.dianji.pangu.sdk.excel.common.SheetConfig;
import com.dianji.pangu.sdk.excel.model.ExcelData;

import java.io.*;
import java.util.List;

/**
 * 常用Excel读写包装，便于快速使用
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/21
 */
@SuppressWarnings("unused")
public class ExcelUtils {

    private ExcelUtils(){
    }

    /**
     * 将实体列表数据写到excel流
     * @param outputStream excel输出流
     * @param entityList 实体列表
     * @param sheetConfig sheet配置
     * @return StringBuilder
     * @throws IOException IO异常
     */
    public static <T> StringBuilder writeToExcel(OutputStream outputStream, List<T> entityList, SheetConfig sheetConfig) throws IOException,ExcelException{
        return writeToExcel(outputStream,entityList,null,sheetConfig);
    }
    /**
     * 将实体列表数据写到excel流
     * @param outputStream excel输出流
     * @param entityList 实体列表
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @return StringBuilder
     * @throws IOException IO异常
     */
    public static <T> StringBuilder writeToExcel(OutputStream outputStream, List<T> entityList,TypeReference<T> typeReference, SheetConfig sheetConfig) throws IOException,ExcelException{
        StringBuilder errorMsg;
        ExcelData writeExcel = ExcelData.open(outputStream);
        try {
            errorMsg = writeExcel.write(entityList, typeReference,sheetConfig);
        }finally{
            writeExcel.close();
        }
        return errorMsg;
    }

    /**
     * 将实体列表数据写到excel文件
     * @param file excel输出文件
     * @param entityList 实体列表
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @return StringBuilder
     * @throws IOException IO异常
     */
    public static <T> StringBuilder writeToExcel(File file, List<T> entityList, TypeReference<T> typeReference, SheetConfig sheetConfig) throws IOException,ExcelException{
        try(OutputStream outputStream =  new FileOutputStream(file)) {
            return writeToExcel(outputStream,entityList,typeReference,sheetConfig);
        }
    }
    /**
     * 将实体列表数据写到excel文件
     * @param file excel输出文件
     * @param entityList 实体列表
     * @param sheetConfig sheet配置
     * @return StringBuilder
     * @throws IOException IO异常
     */
    public static <T> StringBuilder writeToExcel(File file, List<T> entityList, SheetConfig sheetConfig) throws IOException,ExcelException{
        try(OutputStream outputStream =  new FileOutputStream(file)) {
            return writeToExcel(outputStream,entityList,null,sheetConfig);
        }
    }

    /**
     * 将实体列表数据写到excel文件
     * @param filePath excel输出文件路径
     * @param entityList 实体列表
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @return StringBuilder
     * @throws IOException IO异常
     */
    public static <T>  StringBuilder writeToExcel(String  filePath,List<T> entityList, TypeReference<T> typeReference,SheetConfig sheetConfig) throws IOException,ExcelException{
        try(OutputStream outputStream =  new FileOutputStream(filePath)) {
            return writeToExcel(outputStream,entityList,typeReference,sheetConfig);
        }
    }
    /**
     * 将实体列表数据写到excel文件
     * @param filePath excel输出文件路径
     * @param entityList 实体列表
     * @param sheetConfig sheet配置
     * @return StringBuilder
     * @throws IOException IO异常
     */
    public static <T>  StringBuilder writeToExcel(String  filePath,List<T> entityList, SheetConfig sheetConfig) throws IOException,ExcelException{
        try(OutputStream outputStream =  new FileOutputStream(filePath)) {
            return writeToExcel(outputStream,entityList,null,sheetConfig);
        }
    }

    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param inputStream excel输入流
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(InputStream inputStream,Class<T> entityClass,SheetConfig sheetConfig,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        ExcelData readExcel = ExcelData.open(inputStream);
        List<T> entityList;
        try {
            entityList = readExcel.read(entityClass, sheetConfig, errorMsgs);
        }finally {
            readExcel.close();
        }
        return entityList;
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param file excel输入文件
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(File file,Class<T> entityClass,SheetConfig sheetConfig,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        try(InputStream inputStream = new FileInputStream(file)) {
            return readFromExcel(inputStream, entityClass, sheetConfig, errorMsgs);
        }
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param filePath excel输入文件路径
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(String filePath,Class<T> entityClass,SheetConfig sheetConfig,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        try(InputStream inputStream = new FileInputStream(filePath)) {
            return readFromExcel(inputStream, entityClass, sheetConfig, errorMsgs);
        }
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param inputStream excel输入流
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param sheetIndex sheet索引号
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(InputStream inputStream,Class<T> entityClass,SheetConfig sheetConfig, int sheetIndex,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        ExcelData readExcel = ExcelData.open(inputStream);
        List<T> entityList;
        try {
            entityList = readExcel.read(entityClass, sheetConfig, sheetIndex,errorMsgs);
        }finally {
            readExcel.close();
        }
        return entityList;
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param file excel输入文件
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param sheetIndex sheet索引号
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(File file,Class<T> entityClass,SheetConfig sheetConfig,int sheetIndex,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        try(InputStream inputStream = new FileInputStream(file)) {
            return readFromExcel(inputStream, entityClass, sheetConfig, sheetIndex, errorMsgs);
        }
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param filePath excel输入文件路径
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param sheetIndex sheet索引号
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(String filePath,Class<T> entityClass,SheetConfig sheetConfig,int sheetIndex,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        try(InputStream inputStream = new FileInputStream(filePath)) {
            return readFromExcel(inputStream, entityClass, sheetConfig, sheetIndex, errorMsgs);
        }
    }

    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param filePath excel输入文件路径
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @param sheetIndex sheet索引号
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(String filePath,TypeReference<T> typeReference,SheetConfig sheetConfig,int sheetIndex,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        try(InputStream inputStream = new FileInputStream(filePath)) {
            return readFromExcel(inputStream, typeReference, sheetConfig, sheetIndex, errorMsgs);
        }
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param filePath excel输入文件路径
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(String filePath,TypeReference<T> typeReference,SheetConfig sheetConfig,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        try(InputStream inputStream = new FileInputStream(filePath)) {
            return readFromExcel(inputStream, typeReference, sheetConfig, errorMsgs);
        }
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param inputStream excel输入流
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(InputStream inputStream,TypeReference<T> typeReference,SheetConfig sheetConfig,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        ExcelData readExcel = ExcelData.open(inputStream);
        List<T> entityList;
        try {
            entityList = readExcel.read(typeReference, sheetConfig, errorMsgs);
        }finally {
            readExcel.close();
        }
        return entityList;
    }
    /**
     * 从excel文件中读取数据转换成实体列表，读取和实体类注解名一样的sheet，没有则缺省第一sheet
     * @param inputStream excel输入流
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig sheet配置
     * @param sheetIndex sheet索引号
     * @param errorMsgs 存放返回的错误信息列表，null则不返回错误信息
     * @return List 实体列表
     */
    public static <T> List<T> readFromExcel(InputStream inputStream,TypeReference<T> typeReference,SheetConfig sheetConfig, int sheetIndex,List<ErrorMessage> errorMsgs)
            throws IOException,InstantiationException,IllegalAccessException,ExcelException {
        ExcelData readExcel = ExcelData.open(inputStream);
        List<T> entityList;
        try {
            entityList = readExcel.read(typeReference, sheetConfig, sheetIndex,errorMsgs);
        }finally {
            readExcel.close();
        }
        return entityList;
    }

}

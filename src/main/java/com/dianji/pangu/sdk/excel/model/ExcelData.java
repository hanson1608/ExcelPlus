/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: WriteExcel
 * Author:   hanson
 * Date:   2018/5/19 12:05
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.model;

import com.alibaba.fastjson.JSONArray;
import com.dianji.pangu.sdk.excel.common.ErrorMessage;
import com.dianji.pangu.sdk.excel.common.Excel;
import com.dianji.pangu.sdk.excel.common.ExcelException;
import com.dianji.pangu.sdk.excel.common.SheetConfig;
import com.dianji.pangu.sdk.excel.util.TypeReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 写Excel文件
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/19
 */
@SuppressWarnings("unused")
public class ExcelData {
    /**
     * 存放execl对应的workBook
     */
    private Workbook excelWb;
    /**
     * 将wb输出的流
     */
    private OutputStream outputStream = null;
    private InputStream inputStream = null;
    private CellStyle titleCellStyle;
    private CellStyle normalCellStyle;

    /**
     * 创建写Excel 文件实例
     *
     * @param outputStream 输出流对象
     * @return ReadExcel对象
     */
    public static ExcelData open(OutputStream outputStream) {
        ExcelData writeExcel = new ExcelData();
        writeExcel.outputStream = outputStream;
        writeExcel.initWriteExcel();
        return writeExcel;
    }

    /**
     * 创建读Excel 文件实例
     *
     * @param inputStream excel输入流对象,仅支持excel2003以上版本
     * @return ReadExcel对象
     * @throws IOException 流IO异常
     */
    public static ExcelData open(InputStream inputStream) throws ExcelException, IOException {
        ExcelData readExcel = new ExcelData();
        readExcel.excelWb = new XSSFWorkbook(inputStream);
        readExcel.inputStream = inputStream;
        if (readExcel.excelWb.getNumberOfSheets() == 0) {
            throw new ExcelException(ErrorMessage.NO_SHEET_DATA);
        }
        return readExcel;
    }

    /**
     * 初始化读，创建标题样式，标准的单元格样式
     */
    private void initWriteExcel() {
        excelWb = new SXSSFWorkbook();
        //普通单元格样式
        normalCellStyle = excelWb.createCellStyle();
        normalCellStyle.setAlignment(CellStyle.ALIGN_CENTER);

        //标题单元格样式
        titleCellStyle = excelWb.createCellStyle();
        titleCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        // 设置前景色，格式以下方式有效
        ((XSSFCellStyle) titleCellStyle).setFillForegroundColor(new XSSFColor(new java.awt.Color(159, 213, 183)));
        titleCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        Font font = excelWb.createFont();
        font.setColor(HSSFColor.WHITE.index);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        // 设置字体
        titleCellStyle.setFont(font);
    }

    /**
     * 创建sheetData，
     *
     * @param entityClass 实体类
     * @param sheetConfig sheet配置
     * @param sheetName   sheet名字
     * @return sheetData对象
     */
    private <T> SheetData<T> createSheetData(Class<T> entityClass, TypeReference<T> typeReference, SheetConfig sheetConfig, String sheetName) throws ExcelException {
        Sheet sheet;
        SheetData<T> sheetData;

        if (sheetName != null && !"".equals(sheetName)) {
            //避免重复名字
            if (excelWb.getSheetIndex(sheetName) >= 0) {
                sheetName += excelWb.getNumberOfSheets();
            }
            sheet = excelWb.createSheet(sheetName);
        } else {
            sheet = excelWb.createSheet();
        }
        if (typeReference != null) {
            sheetData = new SheetData<>(typeReference, sheet, excelWb);
        } else {
            sheetData = new SheetData<>(entityClass, sheet, excelWb);
        }
        if (sheetConfig != null) {
            sheetData.setSheetConfig(sheetConfig);
        }
        sheetData.setNormalCellStyle(normalCellStyle);
        sheetData.setTitleCellStyle(titleCellStyle);
        //要先设置再初始化
        sheetData.initWrite();
        return sheetData;
    }

    /**
     * 将实体列表写入到新的sheet，每次会新建一个sheet，这个操作只写到workbook，未写到流，close的时候才往流里写。
     *
     * @param entityList    需要写入的实体对象列表
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig   sheet配置
     * @return 错误信息
     */
    @SuppressWarnings("unchecked")
    public <T> StringBuilder write(List<T> entityList, TypeReference<T> typeReference, SheetConfig sheetConfig) throws ExcelException {
        return write(entityList, typeReference, sheetConfig, null);
    }

    /**
     * 将实体列表写入到新的sheet，每次会新建一个sheet，这个操作只写到workbook，未写到流，close的时候才往流里写。
     *
     * @param entityList    需要写入的实体对象列表
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheetConfig   sheet配置
     * @param sheetName     sheet名
     * @return 错误信息
     */
    @SuppressWarnings("unchecked")
    public <T> StringBuilder write(List<T> entityList, TypeReference<T> typeReference, SheetConfig sheetConfig, String sheetName) throws ExcelException {
        StringBuilder errorMsg = new StringBuilder();
        if (entityList.isEmpty()) {
            throw new ExcelException(ErrorMessage.NO_ENTITY_DATA);
        }
        if (entityList.size() >= SheetConfig.SHEET_MAX_ROWS) {
            throw new ExcelException(ErrorMessage.TOO_MANY_ROWS + SheetConfig.SHEET_MAX_ROWS);
        }
        Class<T> entityClass = (Class<T>) entityList.get(0).getClass();
        if (sheetName == null) {
            Excel excelAnnotation = entityClass.getAnnotation(Excel.class);
            if (excelAnnotation != null) {
                sheetName = excelAnnotation.title();
            }
        }
        SheetData<T> sheetData = createSheetData(entityClass, typeReference, sheetConfig, sheetName);
        sheetData.writeSheet(entityList, errorMsg);
        return errorMsg;
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param typeReference 实体类参考类，支持泛型用
     * @param errorMsgs     存放读取过程中的错误信息，不能为null
     * @param sheetConfig   可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @return 返回的实体列表
     */
    public <T> List<T> read(TypeReference<T> typeReference, SheetConfig sheetConfig, List<ErrorMessage> errorMsgs)
            throws InstantiationException, IllegalAccessException, ExcelException {

        SheetData<T> sheetData = makeSheetData(findSheet(typeReference.getRawType()), sheetConfig, null, typeReference);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(errorMsgs);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param typeReference 实体类参考类，支持泛型用
     * @param errorMsgs     存放读取过程中的错误信息，不能为null
     * @param sheetConfig   可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @param sheetIndex    sheet索引号
     * @return 返回的实体列表
     */
    public <T> List<T> read(TypeReference<T> typeReference, SheetConfig sheetConfig, int sheetIndex, List<ErrorMessage> errorMsgs)
            throws InstantiationException, IllegalAccessException, ExcelException {
        Sheet sheet = excelWb.getSheetAt(sheetIndex);
        SheetData<T> sheetData = makeSheetData(sheet, sheetConfig, null, typeReference);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(errorMsgs);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param typeReference 对应的实体类
     * @param errorMsgs     存放读取过程中的错误信息，不能为null
     * @param sheetConfig   可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @return 返回的实体列表
     */
    public <T> List<T> read(TypeReference<T> typeReference, SheetConfig sheetConfig, List<ErrorMessage> errorMsgs, int startRow, int count)
            throws InstantiationException, IllegalAccessException, ExcelException {
        SheetData<T> sheetData = makeSheetData(findSheet(typeReference.getRawType()), sheetConfig, null, typeReference);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(startRow, count, errorMsgs);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param entityClass 对应的实体类
     * @param errorMsgs   存放读取过程中的错误信息，不能为null
     * @param sheetConfig 可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @return 返回的实体列表
     */
    public <T> List<T> read(Class<T> entityClass, SheetConfig sheetConfig, List<ErrorMessage> errorMsgs)
            throws InstantiationException, IllegalAccessException, ExcelException {

        SheetData<T> sheetData = makeSheetData(findSheet(entityClass), sheetConfig, entityClass, null);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(errorMsgs);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param entityClass 对应的实体类
     * @param errorMsgs   存放读取过程中的错误信息，不能为null
     * @param sheetConfig 可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @param sheetIndex  sheet索引号
     * @return 返回的实体列表
     */
    public <T> List<T> read(Class<T> entityClass, SheetConfig sheetConfig, int sheetIndex, List<ErrorMessage> errorMsgs)
            throws InstantiationException, IllegalAccessException, ExcelException {
        Sheet sheet = excelWb.getSheetAt(sheetIndex);
        SheetData<T> sheetData = makeSheetData(sheet, sheetConfig, entityClass, null);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(errorMsgs);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param entityClass 对应的实体类
     * @param errorMsgs   存放读取过程中的错误信息，不能为null
     * @param sheetConfig 可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @return 返回的实体列表
     */
    public <T> List<T> read(Class<T> entityClass, SheetConfig sheetConfig, List<ErrorMessage> errorMsgs, int startRow, int count)
            throws InstantiationException, IllegalAccessException, ExcelException {
        SheetData<T> sheetData = makeSheetData(findSheet(entityClass), sheetConfig, entityClass, null);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(startRow, count, errorMsgs);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param entityClass 对应的实体类
     * @param errorJson   存放读取过程中的错误信息，不能为null
     * @param sheetConfig 可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @param sheetIndex  sheet索引号
     * @return 返回的实体列表
     */
    public <T> List<T> read(Class<T> entityClass, SheetConfig sheetConfig, int sheetIndex, JSONArray errorJson)
            throws InstantiationException, IllegalAccessException, ExcelException {
        Sheet sheet = excelWb.getSheetAt(sheetIndex);
        SheetData<T> sheetData = makeSheetData(sheet, sheetConfig, entityClass, null);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(errorJson);
    }

    /**
     * 根据实体类自动读取所有相关的数据，其他的灵活方式请先构建sheetData,然后再调用相关方法。
     *
     * @param typeReference 实体类参考类，支持泛型用
     * @param errorJson     存放读取过程中的错误信息，不能为null
     * @param sheetConfig   可以为null，null表示使用缺省的配置（第一行为标题，第2行为数据）
     * @param sheetIndex    sheet索引号
     * @return 返回的实体列表
     */
    public <T> List<T> read(TypeReference<T> typeReference, SheetConfig sheetConfig, int sheetIndex, JSONArray errorJson)
            throws InstantiationException, IllegalAccessException, ExcelException {
        Sheet sheet = excelWb.getSheetAt(sheetIndex);
        SheetData<T> sheetData = makeSheetData(sheet, sheetConfig, null, typeReference);
        if (sheetData == null) {
            return new ArrayList<>();
        }
        return sheetData.readToBean(errorJson);
    }

    /**
     * 根据类自动匹配对应的Excel Sheet;匹配规则：先按照类的注解找到对应的sheet，找不到缺省匹配为第一个sheet
     *
     * @param entityClass 实体类
     * @return SheetData
     */
    public <T> SheetData<T> readSheet(Class<T> entityClass, SheetConfig sheetConfig) throws ExcelException {
        return makeSheetData(findSheet(entityClass), sheetConfig, entityClass, null);
    }

    /**
     * 读取Excel 第几个sheet
     *
     * @param sheetIndex  第几个Sheet
     * @param sheetConfig sheet配置
     * @param entityClass 实体类
     * @return SheetData
     */
    public <T> SheetData<T> readSheet(int sheetIndex, SheetConfig sheetConfig, Class<T> entityClass) throws ExcelException {
        Sheet sheet = excelWb.getSheetAt(sheetIndex);
        return makeSheetData(sheet, sheetConfig, entityClass, null);
    }

    /**
     * 根据SheetName读取相应的Excel Sheet;
     *
     * @param sheetName   sheet名
     * @param sheetConfig sheet配置
     * @param entityClass 实体类
     * @return SheetData
     */
    public <T> SheetData<T> readSheet(String sheetName, SheetConfig sheetConfig, Class<T> entityClass) throws ExcelException {
        Sheet sheet = excelWb.getSheet(sheetName);
        return makeSheetData(sheet, sheetConfig, entityClass, null);
    }

    /**
     * 找到和class匹配的sheet，先按照类的注解找到对应的sheet，找不到缺省匹配为第一个sheet
     *
     * @param entityClass 实体类
     * @return Sheet
     */
    private Sheet findSheet(Class<?> entityClass) {
        Sheet sheet;
        Excel excelAnnotation = entityClass.getAnnotation(Excel.class);
        if (excelAnnotation != null && !excelAnnotation.title().isEmpty()) {
            sheet = excelWb.getSheet(excelAnnotation.title());
            if (sheet == null) {
                sheet = excelWb.getSheetAt(0);
            }
        } else {
            sheet = excelWb.getSheetAt(0);
        }
        return sheet;
    }

    /**
     * 构建sheetData,entityClass（非泛型参数类）和typeReference（泛型参数类）二选一
     *
     * @param sheet         sheet对象
     * @param sheetConfig   sheet配置
     * @param entityClass   实体类 泛型采用typeReference时送null
     * @param typeReference 实体类参考类，支持泛型用   不是泛型就送null
     * @return SheetData
     */
    private <T> SheetData<T> makeSheetData(Sheet sheet, SheetConfig sheetConfig, Class<T> entityClass, TypeReference<T> typeReference) throws ExcelException {
        SheetData<T> sheetData;
        if (sheet != null) {
            if (typeReference != null) {
                sheetData = new SheetData<>(typeReference, sheet);
            } else {
                sheetData = new SheetData<>(entityClass, sheet);
            }
            if (sheetConfig != null) {
                sheetData.setSheetConfig(sheetConfig);
            }
            sheetData.initRead();
            return sheetData;
        }
        return null;
    }


    /**
     * 读指定的sheet数据到JSONArray
     *
     * @param sheetIndex  第几个Sheet
     * @param sheetConfig sheet配置
     * @return JSONArray
     * @throws ExcelException 异常
     */
    public JSONArray read(int sheetIndex, SheetConfig sheetConfig) throws ExcelException {
        Sheet sheet = excelWb.getSheetAt(sheetIndex);
        SheetData sheetData = new SheetData(sheetConfig, sheet);
        return sheetData.readToJson();
    }

    /**
     * 关闭资源
     *
     * @throws IOException IO异常
     */
    public void close() throws IOException {
        if (outputStream != null) {
            //这里才真正写到输出流
            excelWb.write(outputStream);
            outputStream.close();
        }
        if (inputStream != null) {
            inputStream.close();
        }
        if (excelWb != null) {
            excelWb.close();
        }
    }

    public int getSheetTotalNumber() {
        return excelWb.getNumberOfSheets();
    }

    public String getSheetName(int sheetIndex) {
        return excelWb.getSheetName(sheetIndex);
    }

}

/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: SheetData
 * Author:   hanson
 * Date:   2018/5/18 21:01
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.model;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.dianji.pangu.sdk.excel.common.*;
import com.dianji.pangu.sdk.excel.util.ExcelReflectUtils;
import com.dianji.pangu.sdk.excel.util.TypeReference;
import com.dianji.pangu.sdk.excel.util.ValidatorUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel的sheet数据
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/18
 */
@SuppressWarnings("unused")
public class SheetData<T> {
    private static final int OTHER_TYPE = 0;
    private static final int INTEGER_TYPE = 1;
    private static final int DECIAML_TYPE = 2;
    private static final int DATE_TYPE = 3;
    /**
     * 存放对应的Excel的Sheet
     */
    private Sheet sheet;
    /**
     * 对当前sheet的读取配置文件
     */
    private SheetConfig sheetConfig;
    /**
     * 实体类
     */
    private Class<T> entityClass;
    /**
     * 类参考，包装解决泛型问题
     */
    private TypeReference<T> typeReference = null;
    /**
     * excel文件标题数组，后面用到，根据索引去标题名称，通过标题名称去字段名称找对应的字段
     */
    private String[] titleNames = null;
    /**
     * 存放当前读取sheet的数据总行数。不包含标题
     */
    private int totalDataRow = 0;
    /**
     * 日期格式
     */
    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    /**
     * 对应实体类的字段树
     */
    private FieldTree entityFieldTree;
    /**
     * 日期缺省单元格样式
     */
    private CellStyle dateCellStyle;
    /**
     * 整数缺省单元格样式
     */
    private CellStyle intCellStyle;
    /**
     * 小数缺省单元格样式
     */
    private CellStyle decimalCellStyle;
    /**
     * 标题缺省单元格样式
     */
    private CellStyle titleCellStyle;
    /**
     * 普通缺省单元格样式
     */
    private CellStyle normalCellStyle;
    /**
     * excel对应的workbook
     */
    private Workbook excelWb;
    /**
     * 每列的单元格样式列表
     */
    private Map<String, CellStyle> cellStyleMap = new HashMap<>();

    protected SheetData(SheetConfig sheetConfig, Sheet sheet) throws ExcelException {
        this.sheet = sheet;
        if (sheetConfig != null) {
            this.sheetConfig = sheetConfig;
        } else {
            this.sheetConfig = new SheetConfig();
        }
        initRead();
    }

    /**
     * 构造方法在包内引用，禁止对外开放，为写构造方法
     *
     * @param entityClass 实体类
     * @param sheet       excel对应的sheet
     */
    SheetData(Class<T> entityClass, Sheet sheet, Workbook wb) {
        this.entityClass = entityClass;
        this.sheetConfig = new SheetConfig();
        this.sheet = sheet;
        this.excelWb = wb;

    }

    /**
     * 构造方法在包内引用，禁止对外开放，为写构造方法
     *
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheet         excel对应的sheet
     */
    SheetData(TypeReference<T> typeReference, Sheet sheet, Workbook wb) {
        this.typeReference = typeReference;
        this.sheetConfig = new SheetConfig();
        this.sheet = sheet;
        this.excelWb = wb;

    }

    /**
     * 构造方法在包内引用，禁止对外开放，为读构造方法
     *
     * @param entityClass 实体类
     * @param sheet       excel对应的sheet
     */
    SheetData(Class<T> entityClass, Sheet sheet) {
        this.entityClass = entityClass;
        this.sheetConfig = new SheetConfig();
        this.sheet = sheet;
    }

    /**
     * 构造方法在包内引用，禁止对外开放，为读构造方法
     *
     * @param typeReference 实体类参考类，支持泛型用
     * @param sheet         excel对应的sheet
     */
    SheetData(TypeReference<T> typeReference, Sheet sheet) {
        this.typeReference = typeReference;
        this.sheetConfig = new SheetConfig();
        this.sheet = sheet;
    }

    /**
     * 初始化写
     */
    public void initWrite() throws ExcelException {
        //根据配置构建样式
        CreationHelper createHelper = excelWb.getCreationHelper();
        //日期样式
        dateCellStyle = excelWb.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(sheetConfig.getDateFormat()));
        dateCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        //整数样式
        intCellStyle = excelWb.createCellStyle();
        intCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(sheetConfig.getIntegerFormat()));
        intCellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        //小数样式
        decimalCellStyle = excelWb.createCellStyle();
        decimalCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(sheetConfig.getDoubleFormat()));
        decimalCellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        // 构建字段映射树
        makeFieldTree();
        //写标题行
        Row titleRow = sheet.createRow(sheetConfig.getTitleRow());
        makeColStyleAndTitle(entityFieldTree, titleRow);
    }

    /**
     * 构建字段映射树
     */
    private void makeFieldTree() throws ExcelException {
        //如果是TypeReference泛型,
        if (typeReference != null) {
            entityFieldTree = ExcelReflectUtils.makeFieldTree(typeReference, sheetConfig.getFieldNameMap());
        } else {
            entityFieldTree = ExcelReflectUtils.makeFieldTree(entityClass, sheetConfig.getFieldNameMap());
        }
    }

    /**
     * 遍历树，构建每个列对应的样式，并填充列标题
     *
     * @param fieldTree 字段映射数
     * @param titleRow  标题行
     */
    private void makeColStyleAndTitle(FieldTree fieldTree, Row titleRow) {
        //是否该字段是基本类型节点
        if (!fieldTree.isBaseClassType()) {
            //遍历所有子节点
            for (FieldTree child : fieldTree.getChildren()) {
                makeColStyleAndTitle(child, titleRow);
            }
        } else {
            //构建对应的列单元格样式
            CellStyle cellStyle = getCellStyleWithType(fieldTree.getField());

            if (!"".equals(fieldTree.getColFormat())) {
                //注解仅设置了列格式,但对齐方式和原来一样。
                CellStyle newCellStyle = excelWb.createCellStyle();
                newCellStyle.setAlignment(cellStyle.getAlignment());
                newCellStyle.setVerticalAlignment(cellStyle.getVerticalAlignment());

                CreationHelper createHelper = excelWb.getCreationHelper();
                newCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(fieldTree.getColFormat()));
                //指向新的样式
                cellStyle = newCellStyle;
            }
            cellStyleMap.put(fieldTree.getColTitleName(), cellStyle);
            // 设置对应列的对应宽带，
            sheet.setColumnWidth(fieldTree.getColIndex(), fieldTree.getColWidth() * sheetConfig.getCellUnitWidth());
            //设置对应列标题单元格
            Cell cell = titleRow.createCell(fieldTree.getColIndex());
            cell.setCellStyle(titleCellStyle);
            cell.setCellValue(fieldTree.getColTitleName());
        }
    }

    /**
     * 根据字段，找到和字段类型匹配的缺省单元格样式
     *
     * @param field 字段
     * @return 返回单元格样式
     */
    private CellStyle getCellStyleWithType(Field field) {
        CellStyle cellStyle;

        int type = getFieldExcelType(field);
        if (type == INTEGER_TYPE) {
            cellStyle = intCellStyle;
        } else if (type == DECIAML_TYPE) {
            cellStyle = decimalCellStyle;
        } else if (type == DATE_TYPE) {
            cellStyle = dateCellStyle;
        } else {
            cellStyle = normalCellStyle;
        }
        return cellStyle;
    }

    /**
     * 将实体列表写入到sheet里，错误信息保存在errorMsg里
     *
     * @param entityList 实体列表
     * @param errorMsg   错误信息
     */
    public void writeSheet(List<T> entityList, StringBuilder errorMsg) {
        int rowIndex = sheetConfig.getDataBeginRow();
        for (T entity : entityList) {
            Row row = sheet.createRow(rowIndex);
            //每个实体填充一行
            beanToRow(entityFieldTree, entity, row, errorMsg);
            rowIndex++;
        }
    }

    /**
     * 遍历字段树，将对象的相应字段填充行
     *
     * @param fieldTree 字段映射数
     * @param obj       对象
     * @param row       标题行
     */
    private void beanToRow(FieldTree fieldTree, Object obj, Row row, StringBuilder errorSb) {
        //当前对象为空则返回，不创建单元格
        if (obj == null) {
            return;
        }
        //根节点单独处理
        if (fieldTree.getField() == null) {
            for (FieldTree child : fieldTree.getChildren()) {
                beanToRow(child, obj, row, errorSb);
            }
            return;
        }
        try {
            //是否该字段是基本类型节点
            if (!fieldTree.isBaseClassType()) {
                Object fieldObj = fieldTree.getField().get(obj);

                //遍历所有子节点
                for (FieldTree child : fieldTree.getChildren()) {
                    //获取当前字段对象，往下寻找该对象的字段
                    beanToRow(child, fieldObj, row, errorSb);
                }
            } else {
                Object value = fieldTree.getField().get(obj);
                //空值则不建cell。
                if (value != null) {
                    //基本类型节点,构建字段树相应的单元格赋值
                    Cell cell = row.createCell(fieldTree.getColIndex());
                    cell.setCellStyle(cellStyleMap.get(fieldTree.getColTitleName()));
                    setCellValue(cell, value);
                    ExcelDataFormatter edf = sheetConfig.getEdf();
                    // 特殊字符的转义
                    if ((edf != null) && (null != edf.get(fieldTree.getColTitleName()))) {
                        cell.setCellValue(edf.get(fieldTree.getColTitleName()).get(value.toString().toLowerCase()));
                    }
                }
            }
        } catch (Exception e) {
            errorSb.append("字段：").append(fieldTree.getField().getName())
                    .append("异常，异常类型：").append(e.getClass());
        }
    }

    /**
     * 初始化读excel文件，构建字段树
     */
    public void initRead() throws ExcelException {
        initTitleName();
        //有效数据行,计算从开始行到数据行之间的有效行数
        int count = 0;
        for (int i = sheet.getFirstRowNum(); i < sheetConfig.getDataBeginRow(); i++) {
            if (sheet.getRow(i) != null) {
                count++;
            }
        }
        //实际数据行数
        totalDataRow = sheet.getPhysicalNumberOfRows() - count;
        //构建字段和excel对应树
        makeFieldTree(titleNames);
    }

    /**
     * 初始化列名
     */
    private void initTitleName() throws ExcelException {
        //判断有没有初始化过，减少重复初始化。
        if (titleNames != null) {
            return;
        }
        Row title = sheet.getRow(sheetConfig.getTitleRow());
        if (title == null) {
            throw new ExcelException(ErrorMessage.NO_TITLE_DATA);
        }
        // 初始化列标题数组
        titleNames = new String[title.getLastCellNum()];

        for (int i = title.getFirstCellNum(); i < title.getLastCellNum(); i++) {
            Cell cell = title.getCell(i);
            if (cell == null) {
                continue;
            }
            if (!cell.getStringCellValue().isEmpty()) {
                titleNames[i] = cell.getStringCellValue();
            }
        }

    }

    /**
     * 构建字段映射树，根据有效字段名过滤（读使用）
     *
     * @param titleNames 有效字段名
     * @throws ExcelException 自定义异常
     */
    private void makeFieldTree(String[] titleNames) throws ExcelException {
        //泛型会使用TypeReference
        if (typeReference != null) {
            entityFieldTree = ExcelReflectUtils.makeFieldTree(typeReference, titleNames, sheetConfig.getFieldNameMap());
        } else if (entityClass != null) {
            entityFieldTree = ExcelReflectUtils.makeFieldTree(entityClass, titleNames, sheetConfig.getFieldNameMap());
        }
    }

    /**
     * 从当前sheet里读取指定数据到JSON数组,
     *
     * @return JSONArray JSON数组
     */
    public JSONArray readToJson() {
        return readToJson(0, totalDataRow);
    }

    /**
     * 从当前sheet里读取指定数据到JSON数组,
     *
     * @param startDataRow 开始行
     * @param count        行数
     * @return JSONArray JSON数组
     */
    public JSONArray readToJson(int startDataRow, int count) {
        Row row;
        int endRow;
        int startRow;
        JSONArray data = new JSONArray();
        //计算起始和结束行
        startRow = startDataRow + sheetConfig.getDataBeginRow();
        endRow = startRow + count;
        if (endRow > sheet.getLastRowNum()) {
            endRow = sheet.getLastRowNum();
        }

        //读取指定行数据到bean数组
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            row = sheet.getRow(rowIndex);
            //判断是不是空行
            if (row == null) {
                break;
            }
            data.add(rowToJson(row));
        }
        return data;
    }

    /**
     * 从当前sheet里读取指定数据到bean列表,
     *
     * @param errorJson JSON格式存放错误信息，null就不保存，将错误行保留下来
     * @return java.util.List
     * @author hanson
     * Date:2018/4/25 21:46
     * @since 1.0.0
     */
    public List<T> readToBean(JSONArray errorJson) throws InstantiationException, IllegalAccessException {
        return readToBean(0, totalDataRow, errorJson);
    }

    /**
     * 从当前sheet里读取指定数据到bean列表,
     *
     * @param errorMsgList 存放错误信息，null就不保存
     * @return java.util.List
     * @author hanson
     * Date:2018/4/25 21:46
     * @since 1.0.0
     */
    public List<T> readToBean(List<ErrorMessage> errorMsgList) throws InstantiationException, IllegalAccessException {
        return readToBean(0, totalDataRow, errorMsgList);
    }

    /**
     * 从当前sheet里读取指定数据到bean列表,
     *
     * @param startDataRow 开始行
     * @param count        行数
     * @param errorMsgList 存放错误信息，null就不保存
     * @return java.util.List
     * @author hanson
     * Date:2018/4/25 21:46
     * @since 1.0.0
     */
    @SuppressWarnings("unchecked")
    public List<T> readToBean(int startDataRow, int count, List<ErrorMessage> errorMsgList)
            throws InstantiationException, IllegalAccessException {
        return readToBean(startDataRow, count, errorMsgList, null);
    }

    /**
     * 从当前sheet里读取指定数据到bean列表,
     *
     * @param startDataRow 开始行
     * @param count        行数
     * @param errorJson    JSON格式存放错误信息，null就不保存，将错误行保留下来
     * @return java.util.List
     */
    public List<T> readToBean(int startDataRow, int count, JSONArray errorJson)
            throws InstantiationException, IllegalAccessException {
        return readToBean(startDataRow, count, null, errorJson);
    }

    /**
     * 从当前sheet里读取指定数据到bean列表,
     *
     * @param startDataRow 开始行
     * @param count        行数
     * @param errorMsgList 存放错误信息，null就不保存
     * @param errorJson    JSON格式存放错误信息，null就不保存，将错误行保留下来
     * @return java.util.List
     * @author hanson
     * Date:2018/4/25 21:46
     * @since 1.0.0
     */
    private List<T> readToBean(int startDataRow, int count, List<ErrorMessage> errorMsgList, JSONArray errorJson)
            throws InstantiationException, IllegalAccessException {
        List<T> entityList = new ArrayList<>();
        Row row;
        int endRow;
        int startRow;

        //计算起始和结束行
        startRow = startDataRow + sheetConfig.getDataBeginRow();
        endRow = startRow + count;
        if (endRow > sheet.getLastRowNum()) {
            endRow = sheet.getLastRowNum();
        }

        //读取指定行数据到bean数组
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            row = sheet.getRow(rowIndex);
            //判断是不是空行
            if (row == null) {
                break;
            }
            T entity = (T) entityFieldTree.getFieldType().newInstance();
            //把该行数据转换成bean，转换出错信息存放在errorsb里
            StringBuilder errorSb = new StringBuilder();
            rowToBean(row, entityFieldTree, entity, errorSb);
            //校验Bean,校验错误信息加入到错误信息中
            validateEntity(entity, errorSb);
            //没有错误就加入到实体列表，否则就忽略
            if (errorSb.length() == 0) {
                entityList.add(entity);
            } else {
                //把错误信息汇总到errorMsg列表里。
                if (errorMsgList != null) {
                    errorMsgList.add(new ErrorMessage(rowIndex, errorSb));
                }

                //把错误信息汇总到errorMsg列表里。
                if (errorJson != null) {
                    JSONObject jsonObject = rowToJson(row);
                    jsonObject.put("错误信息", errorSb);
                    jsonObject.put("错误行", rowIndex);
                    errorJson.add(jsonObject);
                }
            }
        }
        return entityList;
    }

    /**
     * 将行信息转换成JSONObject
     *
     * @param row 行
     * @return 返回对应的JSONObject
     */
    private JSONObject rowToJson(Row row) {
        JSONObject jsonObject = new JSONObject();

        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            //对于超过列名范围的单元格忽略
            if ((i > titleNames.length) || (titleNames[i] == null)) {
                break;
            }
            if (cell == null) {
                jsonObject.put(titleNames[i], "");
            } else {
                jsonObject.put(titleNames[i], getCellValue(cell));
            }
        }
        return jsonObject;

    }

    /**
     * 把该行转换成对应的bean
     *
     * @param row        指明的行
     * @param fieldTree  对应的字段树节点
     * @param currentObj 当前对象
     * @param errorSb    错误信息记录
     */
    private void rowToBean(Row row, FieldTree fieldTree, Object currentObj, StringBuilder errorSb) {
        //根节点，遍历所有子节点
        if (fieldTree.getField() == null) {
            for (FieldTree child : fieldTree.getChildren()) {
                rowToBean(row, child, currentObj, errorSb);
            }
            return;
        }
        //非根节点
        try {
            Field field = fieldTree.getField();

            if (fieldTree.isBaseClassType()) {
                //基本类型就直接赋值
                int index = fieldTree.getColIndex();
                if (index >= 0) {
                    Cell cell = row.getCell(index);
                    if (cell != null) {
                        Object cellValue = getCellValue(cell);
                        if (cellValue != null) {
                            setFieldValue(field, currentObj, cellValue);
                            return;
                        }
                    }
                    //cell==null or cellValue==null
                    field.set(currentObj, null);
                }
                return;
            }
            //如果不是基本类型就往下构建
            //创建字段对象,从树上取，解决泛型问题
            Object fieldObj = fieldTree.getFieldType().newInstance();
            //往下构建字段对象
            for (FieldTree child : fieldTree.getChildren()) {
                rowToBean(row, child, fieldObj, errorSb);
            }
            field.set(currentObj, fieldObj);
            //校验非基本类
            validateEntity(fieldObj, errorSb);
        } catch (Exception e) {
            errorSb.append(String.format("列名：%s，错误:单元格数据读取或转换错误！;", fieldTree.getColTitleName()));
        }
    }

    /**
     * 校验实体对象,校验错误信息加入到错误信息中
     *
     * @param entity  实体对象
     * @param errorSb 存放错误信息的StringBuilder
     */
    private void validateEntity(Object entity, StringBuilder errorSb) {
        List<String> validErrorMsg = ValidatorUtil.validate(entity);
        if (!validErrorMsg.isEmpty()) {
            errorSb.append(ErrorMessage.VERIRY_ERROR);
            for (String str : validErrorMsg) {
                errorSb.append(str).append(",");
            }
        }
    }

    /**
     * 获取当前字段在Excel里的类型，主要区分excel里的特殊类型日期，整数、小数
     *
     * @param field 字段
     * @return int
     * @author hanson
     * Date:2018/4/25 20:58
     * @since 1.0.0
     */
    private int getFieldExcelType(Field field) {
        Class fieldClass = field.getType();
        if (fieldClass.equals(Date.class)) {
            return DATE_TYPE;
        }
        if (fieldClass.equals(Long.class) || fieldClass.equals(Integer.class)
                || fieldClass.equals(long.class) || fieldClass.equals(int.class)) {
            return INTEGER_TYPE;
        }
        if (fieldClass.equals(Float.class) || fieldClass.equals(Double.class) || fieldClass.equals(BigDecimal.class)
                || fieldClass.equals(float.class) || fieldClass.equals(double.class)) {
            return DECIAML_TYPE;
        }
        return OTHER_TYPE;
    }

    /**
     * 设置单元格数据, 根据单元格的数据类型返回缺省的数据格式字符串
     *
     * @param cell  单元格
     * @param value 值
     * @author hanson
     * Date:2018/4/25 21:01
     * @since 1.0.0
     */
    private void setCellValue(Cell cell, Object value) {

        // 如果数据为空，设置一个空值单元格,并结束本次循环,instanceof 包装类同时支持对应的基本类
        if (value == null) {
            cell.setCellValue("");
            return;
        }
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
            return;
        }
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
            return;
        }
        if (value instanceof String) {
            cell.setCellValue((String) value);
            return;
        }
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
            return;
        }
        if (value instanceof Long) {
            cell.setCellValue((Long) value);
            return;
        }
        if (value instanceof Float) {
            cell.setCellValue((Float) value);
            return;
        }
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
            return;
        }
        if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).doubleValue());
            return;
        }
        cell.setCellValue(value.toString());
    }

    /**
     * 获取单元格里的数据
     *
     * @param cell 单元格
     */
    private Object getCellValue(Cell cell) {
        Object value;

        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    value = DateUtil.getJavaDate(cell.getNumericCellValue());
                } else {
                    DecimalFormat df = new DecimalFormat();
                    value = df.format(cell.getNumericCellValue()).replaceAll(",","");
                }
                break;
            case XSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                value = cell.getErrorCellValue();
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
                break;
            case XSSFCell.CELL_TYPE_BLANK:
            default:
                value = null;
                break;
        }
        return value;
    }

    /**
     * 设置目标对象的指定字段值
     *
     * @param field   字段
     * @param destObj 目标对象
     * @param value   值
     * @author hanson
     * Date:2018/4/25 21:52
     * @since 1.0.0
     */
    private void setFieldValue(Field field, Object destObj, Object value)
            throws ParseException, IllegalAccessException {

        String valueStr = value.toString();
        Class fieldClass = field.getType();

        //先判断是否字符串类型，不论是否空都直接赋值
        if (fieldClass.equals(String.class)) {
            field.set(destObj, valueStr);
            return;
        }
        //excel单元格空处理，非String类型都设为null
        if ("".equals(valueStr)) {
            field.set(destObj, null);
            return;
        }
        if (fieldClass.equals(Date.class)) {
            if (value instanceof Date) {
                field.set(destObj, value);
            } else {
                field.set(destObj, sdf.parse(valueStr));
            }
            return;
        }

        if (fieldClass.equals(Float.class) || fieldClass.equals(float.class)) {
            field.set(destObj, new Float(valueStr));
            return;
        }
        if (fieldClass.equals(Double.class) || fieldClass.equals(double.class)) {
            field.set(destObj, new Double(valueStr));
            return;
        }
        if (fieldClass.equals(BigDecimal.class)) {
            field.set(destObj, new BigDecimal(valueStr));
            return;
        }
        //注意：excel数值读入都是小数double,带小数点，对于整数需要进行取整，去掉小数点后面数字
        if (field.getType().equals(Long.class) || fieldClass.equals(long.class)) {
            field.set(destObj, getExcelInteger(valueStr));
            return;
        }
        //只有整型和布尔类型需要判断是否需要数据映射，寻找映射表对应的映射字符串
        valueStr = getFieldTransferMapString(field, valueStr);

        if (fieldClass.equals(Boolean.class) || fieldClass.equals(boolean.class)) {
            if (value instanceof Boolean) {
                field.set(destObj, value);
            } else {
                field.set(destObj, Boolean.parseBoolean(valueStr));
            }
            return;
        }
        if (fieldClass.equals(Integer.class) || fieldClass.equals(int.class)) {
            field.set(destObj, (int) getExcelInteger(valueStr));
        }
    }

    /**
     * 根据field找到特殊映射Map对应的字符串，没有对应的映射，就保持原来的
     *
     * @param field 字段
     * @return Map<String   ,       String>，如果有映射就返回map，否则返回null
     */
    private String getFieldTransferMapString(Field field, String valueStr) {
        Map<String, String> map = null;

        if (sheetConfig.getEdf() != null) {
            Excel excelAnnotation = field.getAnnotation(Excel.class);
            if (excelAnnotation != null) {
                map = sheetConfig.getEdf().get(excelAnnotation.title());
            }
        }
        if (map != null && map.get(valueStr) != null) {
            valueStr = map.get(valueStr);
        }
        return valueStr;
    }

    /**
     * 从excel的数值类型转换成整数，excel数值类型读取的时候都带小数点，需要专门处理
     *
     * @param valueStr 当前的字符串
     * @return long
     */
    private long getExcelInteger(String valueStr) {
        String valueIntStr = valueStr;
        int pos = valueStr.indexOf('.');
        if (pos > 0) {
            valueIntStr = valueStr.substring(0, pos);
        }
        return Long.parseLong(valueIntStr);
    }

    public void setSheetConfig(SheetConfig sheetConfig) {
        this.sheetConfig = sheetConfig;
    }

    public void setTitleCellStyle(CellStyle titleCellStyle) {
        this.titleCellStyle = titleCellStyle;
    }

    public void setNormalCellStyle(CellStyle normalCellStyle) {
        this.normalCellStyle = normalCellStyle;
    }
}

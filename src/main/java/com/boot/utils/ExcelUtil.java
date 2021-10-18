package com.boot.utils;

import com.boot.annonation.ExcelFiled;
import com.boot.constant.Constant;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;

/**
 * @Desc Excel 工具类
 * @Author lixu
 * @Date 2021/10/14 11:26
 */
@Slf4j
@UtilityClass
public class ExcelUtil {

    /**************************************  读取 Excel 数据  **************************************/

    /**
     * 读取 Excel 表中的数据
     */
    public <T> List<T> readExcelToData(MultipartFile file, Class<?> clazz) {

        List<T> list = new ArrayList<T>();

        // 检查传入的文件类型是否是 Excel 类型
        checkFile(file);

        // 创建工作簿
        Workbook workbook = createWorkbook(file);

        // 读取实体中的字段
        Field[] fields = getFields(clazz);

        // 读取数据
        if (Objects.nonNull(workbook)) {

            // 选中 Sheet
            Sheet sheet = workbook.getSheetAt(0);
            if (Objects.isNull(sheet) || sheet.getPhysicalNumberOfRows() == 0) {
                return list;
            }

            // 循环数据
            int firstRowNum = sheet.getFirstRowNum();
            int endRowNum = sheet.getLastRowNum();

            for (int i = firstRowNum; i <= endRowNum; i++) {

                // 选取行
                Row row = sheet.getRow(i);
                if (Objects.isNull(row)) {
                    continue;
                }

                Object object = null;
                try {
                    // 反射实例化对象
                    object = clazz.getDeclaredConstructor().newInstance();
                } catch (Exception e) {
                    throw new RuntimeException("反射实例化对象失败!");
                }

                boolean setValue = false;

                // 循环实体中的各个字段（与外循环相配合）
                for (Field field : fields) {

                    // 获取注解
                    ExcelFiled excelFiled = field.getDeclaredAnnotation(ExcelFiled.class);
                    if (Objects.isNull(excelFiled)) {
                        return null;
                    }

                    // 读取元素
                    Cell cell = row.getCell(excelFiled.columnIndex());

                    // 判断从第几行开始读取
                    if (excelFiled.rowIndex() > i) {
                        break;
                    }

                    // 表示读取到数据了
                    if (!setValue) {
                        setValue = true;
                    }

                    // 转换读取到的数据
                    Object value = getCellValue(cell, field);

                    // 创建对象实例
                    createBean(field, object, value);
                }

                if (setValue) {
                    list.add((T) object);
                }

            }
        }

        return list;

    }

    /**
     * 检查传入的文件类型是否是 Excel 类型
     */
    private void checkFile(MultipartFile file) {
        if (Objects.isNull(file)) {
            throw new RuntimeException("File Is Not Null");
        }

        String fileName = file.getOriginalFilename();

        if (Strings.isEmpty(fileName)) {
            throw new RuntimeException("File Name Is Not Null");
        }

        if (!(fileName.endsWith(Constant.EXCEL_2003) || fileName.endsWith(Constant.EXCEL_2007))) {
            throw new RuntimeException("File is not Excel");
        }
    }

    /**
     * 创建工作簿
     */
    private Workbook createWorkbook(MultipartFile file) {

        Workbook workbook;
        InputStream inputStream;

        try {
            inputStream = file.getInputStream();
            workbook = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            throw new RuntimeException("Workbook Create Failure !");
        }

        return workbook;
    }

    /**
     * 值读取
     */
    private Object getCellValue(Cell cell, Field field) {
        Object cellValue = null;
        if (cell == null) {
            return cellValue;
        }
        // 把数字当成String来读，避免出现1读成1.0的情况
        // 判断数据的类型
        switch (cell.getCellType()) {
            case NUMERIC:
                if (cell.getCellType() == CellType.NUMERIC) {
                    if (DateUtil.isValidExcelDate(cell.getNumericCellValue())) {
                        CellStyle style = cell.getCellStyle();
                        if (style == null) {
                            return false;
                        }
                        int i = style.getDataFormat();
                        String f = style.getDataFormatString();
                        boolean isDate = DateUtil.isADateFormat(i, f);
                        if (isDate) {
                            Date date = cell.getDateCellValue();
                            return cellValue = date;
                        }
                    }
                }
                // 防止科学计数进入
                if (String.valueOf(cell.getNumericCellValue()).toLowerCase().contains("e")) {
                    throw new RuntimeException("excel数据类型错误，请将数字转文本类型！！");
                }
                if ((int) cell.getNumericCellValue() != cell.getNumericCellValue()) {
                    // double 类型
                    cellValue = cell.getNumericCellValue();
                } else {
                    cellValue = (int) cell.getNumericCellValue();
                }
                break;
            // 字符串
            case STRING:
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            // Boolean
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            // 公式
            case FORMULA:
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            // 空值
            case BLANK:
                cellValue = null;
                break;
            // 故障
            case ERROR:
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }

    /**
     * 创建实例
     */
    private <T> void createBean(Field field, T newInstance, Object value) {
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }
        try {
            if (value == null) {
                field.set(newInstance, null);
            } else if (Long.class.equals(field.getType())) {
                field.set(newInstance, Long.valueOf(String.valueOf(value)));
            } else if (String.class.equals(field.getType())) {
                field.set(newInstance, String.valueOf(value));
            } else if (Integer.class.equals(field.getType())) {
                field.set(newInstance, Integer.valueOf(String.valueOf(value)));
            } else if (int.class.equals(field.getType())) {
                field.set(newInstance, Integer.parseInt(String.valueOf(value)));
            } else if (Date.class.equals(field.getType())) {
                field.set(newInstance, (Date) value);
            } else if (Boolean.class.equals(field.getType())) {
                field.set(newInstance, (Boolean) value);
            } else if (Double.class.equals(field.getType())) {
                field.set(newInstance, Double.valueOf(String.valueOf(value)));
            } else if (LocalDate.class.equals(field.getType())) {
                field.set(newInstance, ((Date) value).toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
            } else if (LocalDateTime.class.equals(field.getType())) {
                field.set(newInstance, ((Date) value).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
            } else {
                field.set(newInstance, value);
            }
        } catch (IllegalAccessException e) {
            throw new RuntimeException("excel实体赋值类型转换异常", e);
        }
    }

    /**************************************  导出 Excel 数据  **************************************/

    /**
     * 将数据导出到 Excel
     */
    public <T> byte[] readDataToExcel(List<T> data, Class<T> clazz) {

        // 创建 Workbook
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        // 获取字段列表
        Field[] fields = getFields(clazz);

        // 定义变量，存储标题列表
        List<String> headerList = new ArrayList<>();

        // 定义变量，定义与标题列表对应的字段 Name
        List<String> variables = new ArrayList<>();

        // 定义变量 行号
        int rowIndex = 0;

        // 创建 Sheet
        Sheet sheet = workbook.createSheet();

        // 创建表头行
        Row headerRow = sheet.createRow(rowIndex++);

        // 表头处理(先根据字段 sort 排序，再录入文字)
        int fieldsLength = fields.length;
        Field[] fieldsSorted = fieldSort(fields);

        for (int i = 0; i < fieldsLength; i++) {
            Field field = fieldsSorted[i];
            ExcelFiled excelFiled = field.getAnnotation(ExcelFiled.class);
            String header = excelFiled.title();

            // 加入标题列
            headerList.add(header);
            headerRow.createCell(i).setCellValue(header);

            // 加入字段列
            variables.add(field.getName());
        }

        // 数据处理
        int size = data.size();
        for (int i = 0; i < size; i++) {

            // 创建行
            Row row = sheet.createRow(rowIndex + i);

            // 取一行数据
            T rowData = data.get(i);
            Class<?> clz = rowData.getClass();

            // 填充列
            int columnSize = variables.size();
            for (int j = 0; j < columnSize; j++) {
                Field field;
                try {
                    field = clz.getDeclaredField(variables.get(j));
                } catch (Exception e) {
                    throw new RuntimeException("Field Read Failure!");
                }
                field.setAccessible(true);
                String key = field.getName();
                Object val = null;

                try {
                    val = field.get(rowData);
                } catch (Exception e) {
                    throw new RuntimeException("Val Get Failure！");
                }

                row.createCell(j).setCellValue(String.valueOf(val));
            }
        }


        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            workbook.write(bos);
            return bos.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        } finally {
            try {
                if (Objects.nonNull(workbook)) {
                    workbook.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e.getMessage());
            }
        }
    }

    /**
     * 将 Field[] 根据 sort 排序
     */
    private Field[] fieldSort(Field[] fields) {
        Arrays.parallelSort(fields, (field1, field2) -> {
            ExcelFiled excelFiled1 = field1.getAnnotation(ExcelFiled.class);
            ExcelFiled excelFiled2 = field2.getAnnotation(ExcelFiled.class);
            int sort1 = excelFiled1.sort();
            int sort2 = excelFiled2.sort();

            if (sort1 > sort2) {
                return 1;
            } else if (sort1 < sort2) {
                return -1;
            }
            return 0;
        });

        return fields;
    }

    /**************************************     公用方法      **************************************/
    /**
     * 获取字段列
     */
    private <T> Field[] getFields(Class<T> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        List<Field> tempList = new ArrayList<>();
        if (Objects.isNull(fields) || fields.length == 0) {
            throw new RuntimeException("Please Define Entity Column");
        }

        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelFiled.class)) {
                tempList.add(field);
            }
        }

        Field[] res = new Field[tempList.size()];
        tempList.toArray(res);
        return res;
    }
}

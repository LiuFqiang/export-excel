package utils;

import annotation.Excel;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.ListUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.apache.commons.lang3.StringUtils.isNotBlank;

public class ExportExcelUtil {

    private static HSSFWorkbook workbook;

    private static HSSFCellStyle leftStyle;

    private static HSSFCellStyle rightStyle;

    private static HSSFCellStyle centerStyle;

    private static HSSFCellStyle defaultTitleStyle;
    //取色板
    private static HSSFPalette palette;

    private static final int sheetSize = 65535;

    private static Map<String, Object> codeMap = new HashMap<>();

    /***
     *
     * @param list      导出数据集合
     * @param sheetName 工作表名称
     * @param clazz     实体对象
     * @return
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws IOException
     */
    public static InputStream export(List<?> list, String sheetName, Class<?> clazz) throws InvocationTargetException, IllegalAccessException, IOException {

        if (list == null) {
            list = new ArrayList<>();
        }

        handleSheetName(sheetName);

        workbook = new HSSFWorkbook();
        leftStyle = workbook.createCellStyle();
        rightStyle = workbook.createCellStyle();
        centerStyle = workbook.createCellStyle();
        defaultTitleStyle = workbook.createCellStyle();
        palette = workbook.getCustomPalette();

        Field[] fields = clazz.getDeclaredFields();

        List<Field> fieldsList = Arrays.stream(fields).filter(f -> f.isAnnotationPresent(Excel.class)).sorted(Comparator.comparing(c -> {
            Excel annotation = c.getAnnotation(Excel.class);
            return annotation.sort();
        })).collect(Collectors.toList());

        Map<String, Method> methods = ExportExcelUtil.getMethod(clazz, fieldsList, "get");

        List<? extends List<?>> partitionList = ListUtils.partition(list, sheetSize);
        if (CollectionUtils.isEmpty(partitionList)) {
            HSSFSheet sheet = workbook.createSheet(sheetName);
            createTitleRow(fieldsList, sheet);
        }

        for (int sheetNum = 0; sheetNum < partitionList.size(); sheetNum ++) {
            String presentName = sheetName + (sheetNum + 1);
            if (partitionList.size() == 1) {
                presentName = sheetName;
            }
            HSSFSheet sheet = workbook.createSheet(presentName);

            // 创建标题行
            createTitleRow(fieldsList, sheet);

            // 创建普通行
            for (int i = 0; i < partitionList.get(sheetNum).size(); i++) {
                Row sheetRow = sheet.createRow(i + 1);
                sheetRow.setHeight((short) (16 * 20));
                Object targetObj = partitionList.get(sheetNum).get(i);
                int cellIndex = 0;
                for (int j = 0; j < fieldsList.size(); j++, ++cellIndex) {
                    Field field = fieldsList.get(j);
                    Method method = methods.get(field.getName());
                    Object value = method.invoke(targetObj);
                    Excel annotation = field.getAnnotation(Excel.class);

                    // 处理默认值
                    if (null == value) {
                        if (isNotBlank(annotation.defaultValue())) {
                            value = annotation.defaultValue();
                        } else {
                            value = "";
                        }
                    }

                    // 转换日期格式
                    if (isNotBlank(annotation.pattern()) && field.getType().equals(Date.class) && isNotBlank(String.valueOf(value))) {

                    }

                    // 码值转换
                    if (annotation.codeValue().length > 0 && isNotBlank(String.valueOf(value))) {
                        value = replaceCode(annotation.codeValue(), String.valueOf(value), field.getName());
                    }

                    ExportExcelUtil.setCellValues(sheetRow, cellIndex, value, annotation);
                }
            }
        }

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        os.flush();
        os.close();
        workbook.close();
        return new ByteArrayInputStream(os.toByteArray());
    }

    private static void createTitleRow(List<Field> fieldsList, HSSFSheet sheet) {
        Row titleRow = sheet.createRow(0);
        for (int i = 0; i < fieldsList.size(); i++) {
            Excel annotation = fieldsList.get(i).getAnnotation(Excel.class);
            ExportExcelUtil.setCellValues(titleRow, i, annotation.name(), annotation);

            // 设置列宽
            double width = annotation.width();
            if (width * 256 > 65280) {
                width = 255;
            }
            titleRow.setHeightInPoints(18);
            titleRow.setHeight((short) (18 * 20));
            sheet.setColumnWidth(i, (int) (width * 256));
        }
    }

    private static void setCellValues(Row row, Integer currentIndex, Object value, Excel annotation) {
        Cell cell = row.createCell(currentIndex);

        // 第一行
        if (row.getRowNum() == 0) {
            cell.setCellStyle(getTitleDefaultStyle());
        } else {
            // 默认样式
            if (annotation.align() == Excel.Align.LEFT) {
                cell.setCellStyle(getAlignLeftCellStyle());
            } else if (annotation.align() == Excel.Align.RIGHT) {
                cell.setCellStyle(getAlignRightCellStyle());
            } else {
                cell.setCellStyle(getDefaultCellStyle());
            }
        }

        cell.setCellValue(value.toString());
    }

    private static Map<String, Method> getMethod(Class<?> clazz, List<Field> fieldsList, String methodKeyWord) {
        Map<String, Method> methodMap = new HashMap<String, Method>();
        /* 获取类模板所有属性 */
        for (Field value : fieldsList) {
            /* 获取属性名并组装方法名称 */
            String fieldName = value.getName();
            /*
             * methodKeyWord = 'get'
             * fieldName = 'name'
             * fieldName.substring(0, 1).toUpperCase() = 'N'
             * fieldName.substring(1) = 'ame()'
             * getMethodName = 'getName()'
             * */
            String getMethodName = methodKeyWord +
                    fieldName.substring(0, 1).toUpperCase() +
                    fieldName.substring(1);
            try {
                Method method = clazz.getMethod(getMethodName);
                /*
                 * 存储内容为: id,getId();
                 * name,getName();
                 * */
                methodMap.put(fieldName, method);
            } catch (NoSuchMethodException e) {
                e.printStackTrace();
            }
        }
        return methodMap;
    }

    /**
     * 码值替换
     */
    private static String replaceCode(String[] codeValue, String value, String field) {
        if (!codeMap.containsKey(field)) {
            List<Map<String, String>> codeList = Arrays.stream(codeValue).filter(f -> isNotBlank(f) && f.contains("_")).map(code -> {
                String[] codeAndValue = code.split("_");
                Map<String, String> codeMap = new HashMap<String, String>() {{
                    put(codeAndValue[0], codeAndValue[1]);
                }};
                return codeMap;
            }).collect(Collectors.toList());
            codeMap.put(field, codeList);
        }

        List<Map<String, String>> currCodeList = (List<Map<String, String>>) codeMap.get(field);
        for (Map<String, String> codeMap : currCodeList) {
            if (codeMap.containsKey(value)) {
                return codeMap.get(value);
            }
        }
        return value;
    }

    private static String handleSheetName(String sheetName) {
        if (isBlank(sheetName)) {
            sheetName = "Sheet";
        }
        sheetName = sheetName.replaceAll("[\\s\\\\/:\\*\\?\\\"<>\\|]", "");
        if (sheetName.length() > 30) {
            sheetName = sheetName.substring(0, 30);
        }
        return sheetName;
    }

    /**
     * 默认的cell样式，垂直居中
     *
     * @return
     */
    private static HSSFCellStyle setBaseCellStyle(HSSFCellStyle style) {
        style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
//        style.setBorderBottom(BorderStyle.THIN); //下边框
//        style.setBorderLeft(BorderStyle.THIN);//左边框
//        style.setBorderTop(BorderStyle.THIN);//上边框
//        style.setBorderRight(BorderStyle.THIN);//右边框
        return style;
    }

    private static HSSFCellStyle getDefaultCellStyle() {
        centerStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        setBaseCellStyle(centerStyle);
        return centerStyle;
    }

    private static HSSFCellStyle getAlignLeftCellStyle() {
        leftStyle.setAlignment(HorizontalAlignment.LEFT);//水平居中
        leftStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        setBaseCellStyle(leftStyle);
        return leftStyle;
    }

    private static HSSFCellStyle getAlignRightCellStyle() {
        rightStyle.setAlignment(HorizontalAlignment.RIGHT);//水平居中
        rightStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        setBaseCellStyle(rightStyle);
        return rightStyle;
    }

    /**
     * 默认标题样式
     */
    private static HSSFCellStyle getTitleDefaultStyle() {
        Font f = workbook.createFont();
        f.setFontHeightInPoints((short) 11);
        f.setBold(true);// 加粗
        f.setFontName("宋体");
        f.setColor((short) 9);
        defaultTitleStyle.setFont(f);
        defaultTitleStyle.setBorderBottom(BorderStyle.THIN); //下边框
        defaultTitleStyle.setBorderLeft(BorderStyle.THIN);//左边框
        defaultTitleStyle.setBorderTop(BorderStyle.THIN);//上边框
        defaultTitleStyle.setBorderRight(BorderStyle.THIN);//右边框

        //设置颜色
        HSSFColor hssfColor = palette.findSimilarColor(140, 130, 130);
        defaultTitleStyle.setFillForegroundColor(hssfColor.getIndex());
        defaultTitleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        defaultTitleStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        defaultTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        setBaseCellStyle(defaultTitleStyle);
        return defaultTitleStyle;
    }
}

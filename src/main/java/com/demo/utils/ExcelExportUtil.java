package com.demo.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @className: ExcelExportUtil
 * @package: com.demo.utils
 * @describe: Excel文件的导出
 * @auther: liuzhiyong
 * @date: 2018/7/16
 * @time: 下午 2:54
 */
public class ExcelExportUtil {

    public static final Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);

    /**
     * @methodName: doExportExcel
     * @param: sheetName 工作表的名称
     *          titleName 表头
     *          headers   列表名
     *          dataSet   内容
     *          resultUrl 导出的位置
     *          pattern   时间类型的数据格式
     * @describe: 导出Excel
     * @auther: liuzhiyong
     * @date: 2018/7/16
     * @time: 下午 2:55
     */
    public static void doExportExcel(String sheetName, String titleName, String[] headers, Collection<?> dataSet, String resultUrl, String pattern){
        logger.info("-------------------导出数据开始-------------------");
        //声明一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //声明一个指定名字的工作表
        HSSFSheet sheet = workbook.createSheet(sheetName);
        //设置表格的默认宽度
        sheet.setDefaultColumnWidth(20);
        //设置表格首行合并并居中
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,headers.length-1));
        //设置标题样式
        HSSFCellStyle titleStyle = createTitleStyleForExport(workbook);
        //设置工作簿的样式
        HSSFCellStyle headerStyle = createHeaderStyleForExport(workbook);
        //设置表中数据的格式
        HSSFCellStyle dataCellStyle = createDataCellStyleForExport(workbook);
        //创建标题行-设置格式-赋值
        HSSFRow titleRow = sheet.createRow(0);
        HSSFCell titleCell = titleRow.createCell(0);
        titleCell.setCellStyle(titleStyle);
        titleCell.setCellValue(titleName);
        //创建列首-设置格式-赋值
        HSSFRow row = sheet.createRow(1);
        for(int i = 0;i < headers.length;i++){
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(headerStyle);
            HSSFRichTextString textValue = new HSSFRichTextString(headers[i]);
            cell.setCellValue(textValue);
        }

        Iterator<?> it = dataSet.iterator();
        int index = 1;
        while(it.hasNext()){
            index++;
            row = sheet.createRow(index);
            Object obj = it.next();
            //通过反射获取属性,动态的调用getXxx方法获取属性值
            Field[] fields = obj.getClass().getDeclaredFields();
            for(int i = 0;i< fields.length;i++){
                HSSFCell cell = row.createCell(i);
                cell.setCellStyle(dataCellStyle);
                Field field = fields[i];
                String fieldName = field.getName();
                //方法名
                String getMethodName = "get" + fieldName.substring(0,1).toUpperCase() + fieldName.substring(1);
                try {
                    Class<?> aClass = obj.getClass();
                    Method method = aClass.getMethod(getMethodName, new Class[]{});
                    Object value = method.invoke(obj, new Object[]{});
                    String textValue = null;
                    //如果属性是日期格式，按照指定样式装换
                    if(value instanceof Date){
                        Date date = (Date) value;
                        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                        textValue = sdf.format(date);
                    }else {
                        //其他的统一当做字符串处理
                        textValue = value.toString();
                    }

                    //判断属性是否为空
                    if(textValue != null){
                        Pattern p = Pattern.compile("^\\d+(\\.\\d+)?$");
                        Matcher matcher = p.matcher(textValue);
                        if (matcher.matches()) {
                            // 是数字当作double处理
                            cell.setCellValue(Double.parseDouble(textValue));
                        } else {
                            // 不是数字做普通处理
                            cell.setCellValue(textValue);
                        }
                    }
                    OutputStream out = null;
                    try{
                        out = new FileOutputStream(resultUrl);
                        workbook.write(out);
                    }catch (IOException e){
                        e.printStackTrace();
                    }finally {
                        if(out != null){
                            out.close();
                        }
                    }
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                } catch (IllegalArgumentException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
        logger.info("-------------------导出数据结束-------------------");
    }

    /**
     * @methodName: createTitleStyleForExport
     * @param: [workbook 工作簿]
     * @describe: 设置标题的样式
     * @auther: liuzhiyong
     * @date: 2018/7/16
     * @time: 下午 3:16
     */
    private static HSSFCellStyle createTitleStyleForExport(HSSFWorkbook workbook){
        logger.info("-------------------设置标题格式-------------------");
        //声明[标题]样式,并设置[标题]样式
        HSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
        titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        titleStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        titleStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        titleStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        titleStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        //声明[标题]字体，并设置[标题]字体
        HSSFFont titleFont = workbook.createFont();
        titleFont.setColor(HSSFColor.WHITE.index);
        titleFont.setFontHeightInPoints((short) 24);
        titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        //将[标题字体]应用到[标题样式]
        titleStyle.setFont(titleFont);

        return titleStyle;
    }

    /**
     * @methodName: createHeaderStyleForExport
     * @param: [workbook 工作簿]
     * @describe: 设置列首的样式
     * @auther: liuzhiyong
     * @date: 2018/7/16
     * @time: 下午 3:24
     */
    private static HSSFCellStyle createHeaderStyleForExport(HSSFWorkbook workbook){
        logger.info("-------------------设置列首样式-------------------");
        //声明[列首]样式，并设置[列首]样式
        HSSFCellStyle headersStyle = workbook.createCellStyle();
        headersStyle.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
        headersStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        headersStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        headersStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        headersStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        headersStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        headersStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        //声明[列首]字体，并设置[列首]字体
        HSSFFont headersFont = workbook.createFont();
        headersFont.setColor(HSSFColor.VIOLET.index);
        headersFont.setFontHeightInPoints((short) 12);
        headersFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        //将[列首字体]应用到[列首样式]
        headersStyle.setFont(headersFont);

        return headersStyle;
    }

    /**
     * @methodName: createDataCellStyleForExport
     * @param: [workbook 工作薄]
     * @describe: 设置表中数据的样式
     * @auther: liuzhiyong
     * @date: 2018/7/16
     * @time: 下午 3:24
     */
    private static HSSFCellStyle createDataCellStyleForExport(HSSFWorkbook workbook){
        logger.info("-------------------设置表中数据样式-------------------");
        //声明[表中数据]样式，并设置[表中数据]样式
        HSSFCellStyle dataSetStyle = workbook.createCellStyle();
        dataSetStyle.setFillForegroundColor(HSSFColor.GOLD.index);
        dataSetStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        dataSetStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        dataSetStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        dataSetStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        dataSetStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        dataSetStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        dataSetStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        //声明[表中数据]字体，并设置[表中数据]字体
        HSSFFont dataSetFont = workbook.createFont();
        dataSetFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        dataSetFont.setColor(HSSFColor.BLUE.index);

        //将[表中数据字体]应用到[表中数据样式]
        dataSetStyle.setFont(dataSetFont);

        return dataSetStyle;
    }

}

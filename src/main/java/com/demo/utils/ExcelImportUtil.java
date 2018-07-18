package com.demo.utils;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @className: ExcelImportUtil
 * @package: com.demo.utils
 * @describe: Excel文件的导入
 * @auther: liuzhiyong
 * @date: 2018/7/16
 * @time: 下午 5:35
 */
public class ExcelImportUtil {

    public static final Logger logger = LoggerFactory.getLogger(ExcelImportUtil.class);

    //正则表达式，用于匹配属性的首字母
    public static final String REGEX = "[a-zA-Z]";

    /**
     * @methodName: doImportExcel
     * @param: originUrl 文件的位置
     *          startRow 起始行(从哪一行开始)
     *          endRow   结束行(到哪一行结束
     *                          (0表示所有行;
     *                          正数表示到第几行结束;
     *                          负数表示到倒数第几行结束)
     *                          )
     *          clazz 要返回的对象集合的类型
     * @describe: Excel文件的导入
     * @auther: liuzhiyong
     * @date: 2018/7/16
     * @time: 下午 5:38
     */
    private static List<Object> doImportExcel(String originUrl, int startRow, int endRow, Class<?> clazz) throws IOException {
        File file = new File(originUrl);
        if(!file.exists()){
            logger.info(file.getName() + " 文件不存在!");
            throw new IOException("文件名为 " + file.getName() + "的Excel文件不存在!");
        }
        HSSFWorkbook workbook = null;
        FileInputStream fis = null;
        //Excel中行的集合
        List<Row> rowList = new ArrayList<Row>();
        try {
            //创建文件输入流
            fis = new FileInputStream(file);
            //通过文件输入流去读取文件内容
            workbook = new HSSFWorkbook(fis);
            //获取文件中的第一个sheet表
            HSSFSheet sheet = workbook.getSheetAt(0);
            //获取sheet表中的最后一行
            int rowNum = sheet.getLastRowNum();
            if(rowNum <= 0){
                logger.info( "{} 表的内容为空！",sheet.getSheetName());
                throw new IOException(sheet.getSheetName() + " 表的内容为空！");
            }
            for(int i = startRow;i <= rowNum + endRow;i++){
                Row row = sheet.getRow(i);
                if(row != null){
                    //将每一行加入到list中
                    rowList.add(row);
                    for(int j = 0;j < row.getLastCellNum();j++){
                        String value = getCellValue(row.getCell(j));
                        if(!"".equals(value)){
                            logger.info("开始读取单元格内容！");
                        }
                    }
                }
            }
        }catch (IOException e) {
            e.printStackTrace();
            logger.error("错误信息: {}",e.getMessage());
            throw new IOException("错误信息: 文件读取失败！" + e.getMessage());
        }
        return returnObjectList(rowList,clazz);
    }

    /**
     * @methodName: getCellValue
     * @param: [cell 单元格]
     * @describe: 获取单元格的值
     * @auther: liuzhiyong
     * @date: 2018/7/16
     * @time: 下午 5:43
     */
    private static String getCellValue(Cell cell){
        Object result = null;
        if(cell != null){
            switch (cell.getCellType()){
                case Cell.CELL_TYPE_STRING :
                    result = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    result = cell.getNumericCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    result = cell.getCellFormula();
                    break;
                case Cell.CELL_TYPE_ERROR:
                    result = cell.getErrorCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                default:
                    break;
            }
        }
        return result.toString();
    }

    /**
     * @methodName: returnObjectList
     * @param: rowList 行的列表
     *          clazz   要返回的对象集合的类型
     * @describe: 返回指定的对象集合
     * @auther: liuzhiyong
     * @date: 2018/7/18
     * @time: 上午 10:42
     */
    public static List<Object> returnObjectList(List<Row> rowList,Class<?> clazz){
        List<Object> objectList = null;
        Object object;
        String attribute = null;
        String value = null;
        int j = 0;
        try {
            objectList = new ArrayList<Object>();
            //获取属性
            Field[] fields = clazz.getDeclaredFields();
            if(!CollectionUtils.isEmpty(rowList)){
                for(Row row:rowList){
                    j = 0;
                    object = clazz.newInstance();
                    for(Field field:fields){
                        //获取属性名
                        attribute = field.getName().toString();
                        //获取对应单元格的内容
                        value = getCellValue(row.createCell(j));
                        //给指定的属性赋值
                        setAttributeValue(object,attribute,value);
                        j++;
                    }
                    objectList.add(object);
                }
            }
        }catch (Exception e){
            e.printStackTrace();
            logger.info("异常信息详情：{}",e.getMessage());
        }

        return objectList;
    }

    /**
     * @methodName: setAttributeValue
     * @param: object     指定对象
     *          attribute  指定属相
     *          value      值
     * @describe: 给指定的对象的指定属性赋值
     * @auther: liuzhiyong
     * @date: 2018/7/18
     * @time: 上午 10:57
     */
    public static void setAttributeValue(Object object,String attribute,String value){
        logger.info("---------------------为属性赋值---------------------");
        //获取该属性对应的get/set方法名
        String method_name = convertToMethodName(attribute, object.getClass(), true);
        //获取方法数组
        Method[] methods = object.getClass().getMethods();
        /*
            在java中确定唯一方法的办法是方法名(参数)，但是我们通过get/set加上属性名来获取方法，由于属性名不允许重复，那理论上方法名也不会重复
         */
        for(Method method:methods)
            try {
                if (method.getName().equals(method_name)) {
                    Class<?>[] parameterTypes = method.getParameterTypes();
                    Class<?> type = parameterTypes[0];
                    if (type == int.class || type == Integer.class) {
                        method.invoke(object, Integer.valueOf(value));
                        break;
                    } else if (type == float.class || type == Float.class) {
                        method.invoke(object, Float.valueOf(value));
                        break;
                    } else if (type == double.class || type == Double.class) {
                        method.invoke(object, Double.valueOf(value));
                        break;
                    } else if (type == byte.class || type == Byte.class) {
                        method.invoke(object, Byte.valueOf(value));
                        break;
                    } else if (type == boolean.class || type == Boolean.class) {
                        method.invoke(object, Boolean.valueOf(value));
                        break;
                    } else if (type == Date.class) {
                        try {
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            Date date = sdf.parse(value);
                            method.invoke(object,date);
                        }catch (ParseException e) {
                            logger.info("日期解析错误");
                        }
                        break;
                    }else {
                        method.invoke(object,type.cast(value));
                        break;
                    }
                }
            }  catch (IllegalArgumentException e) {
                logger.info("参数异常，详细信息：{}",e.getMessage());
            } catch (IllegalAccessException e) {
                logger.info("安全权限异常,反射调用的是private方法。详细信息：{}",e.getMessage());
            } catch (InvocationTargetException e) {
                e.printStackTrace();
                logger.info("反射异常，详细信息：{}",e.getMessage());
            }
    }

    /**
     * @methodName: convertToMethodName
     * @param:  attribute   属性名
     *           clazz,      对象
     *           isSet       是否是set方法
     * @describe: 根据属性值生成对应的get/set方法
     * @auther: liuzhiyong
     * @date: 2018/7/18
     * @time: 上午 11:01
     */
    public static String convertToMethodName(String attribute,Class<?> clazz,boolean isSet) {
        //创建表达式
        Pattern pattern = Pattern.compile(REGEX);
        //创建属性匹配规则
        Matcher matcher = pattern.matcher(attribute);
        StringBuffer buffer = new StringBuffer();
        //如果前缀是set开头，则设置set
        if(isSet){
            buffer.append("set");
        }else {
            try {
                //根据属性名获取该属性的数据类型
                Field field = clazz.getDeclaredField(attribute);
                //若该属性属于布尔类型，则设置is开头，否则设置get开头
                if(field.getType() == boolean.class || field.getType() == Boolean.class){
                    buffer.append("is");
                }else {
                    buffer.append("get");
                }
            }catch (NoSuchFieldException e){
                e.printStackTrace();
                logger.info("未找到属性类型，具体信息：{}",e.getMessage());
            }
        }
        //针对属性名是非"_"开始的属性，将首字母大写
        if(attribute.charAt(0) != '_' || matcher.find()){
            buffer.append(
                    matcher.replaceFirst(
                            matcher.group(attribute).toUpperCase()
                    )
            );
        }else {
            buffer.append(attribute);
        }

        return buffer.toString();
    }
}

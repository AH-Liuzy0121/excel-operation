package com.demo.utils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Method;
import java.util.List;

/**
 * @className: CSVUtil
 * @package: com.jy.modules.loan.advice
 * @describe: csv文件操作
 * @auther: liuzhiyong
 * @date: 2018/7/30
 * @time: 下午 5:18
 */
public class CSVUtil<T> {

    /**
     * 生成为CVS文件
     *
     * @param titles     csv文件的列表头(表头信息)
     * @param fileds     导出队形的属性数组
     * @param exportData 源数据List
     * @param outPutPath 文件路径
     * @param fileName   文件名称
     * @return
     */
    @SuppressWarnings("rawtypes")
    public File createCSVFile(String[] titles, String[] fileds, List<T> exportData, String outPutPath, String fileName) {
        File csvFile = null;
        BufferedWriter csvFileOutputStream = null;
        try {
            File file = new File(outPutPath);
            if (!file.exists()) {
                file.mkdir();
            }

            csvFile = new File(file+ "/" + fileName +".csv");
            System.out.println("csvFile：" + csvFile);
            // UTF-8使正确读取分隔符","
            csvFileOutputStream = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), "GBK"), 1024);
            // 写入文件头部
            for (String title : titles) {
                csvFileOutputStream.write(title);
                csvFileOutputStream.write(",");
            }

            csvFileOutputStream.write("\r\n");
            // 写入文件内容,
            for (int j = 0; exportData != null && !exportData.isEmpty() && j < exportData.size(); j++) {
                T t = (T) exportData.get(j);
                Class clazz = t.getClass();
                String[] contents = new String[fileds.length];
                for (int i = 0; fileds != null && i < fileds.length; i++) {
                    String filedName = toUpperCaseFirstOne(fileds[i]);
                    Method method = clazz.getMethod(filedName);
                    method.setAccessible(true);
                    Object obj = method.invoke(t);
                    String str = String.valueOf(obj);
                    if (str == null || str.equals("null"))
                        str = "";
                    contents[i] = str;

                }

                for (int n = 0; n < contents.length; n++) {
                    // 将生成的单元格添加到工作表中
                    csvFileOutputStream.write(contents[n] + "\t");//此处加上"\t"是为了避免出现科学计数法
                    csvFileOutputStream.write(",");

                }
                csvFileOutputStream.write("\r\n");
            }
            csvFileOutputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if(csvFileOutputStream != null){
                    csvFileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return csvFile;
    }

    /**
     * 将第一个字母转换为大写字母并和get拼合成方法
     *
     * @param origin
     * @return
     */
    private static String toUpperCaseFirstOne(String origin) {
        StringBuffer sb = new StringBuffer(origin);
        sb.setCharAt(0, Character.toUpperCase(sb.charAt(0)));
        sb.insert(0, "get");
        return sb.toString();
    }

    /**
     * @methodName: exportFile
     * @param: [response, csvFilePath, fileName]
     * @describe: 下载文件
     * @auther: liuzhiyong
     * @date: 2018/8/28
     * @time: 下午 3:15
     */
    public static void exportFile(HttpServletResponse response, String csvFilePath, String fileName) throws IOException {
        response.setContentType("application/csv;charset=GBK");
        response.setHeader("Content-Disposition", "attachment;  filename="
                + new String(fileName.getBytes("GBK"), "ISO8859-1"));

        InputStream in = null;
        try {
            in = new FileInputStream(csvFilePath);
            int len = 0;
            byte[] buffer = new byte[1024];
            response.setCharacterEncoding("GBK");
            OutputStream out = response.getOutputStream();
            while ((len = in.read(buffer)) > 0) {
                out.write(buffer, 0, len);
            }
        } catch (FileNotFoundException e) {
            System.out.println(e);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }
}

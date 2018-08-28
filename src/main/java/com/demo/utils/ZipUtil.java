package com.demo.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @className: ZipUtil
 * @package: com.demo.utils
 * @describe: 文件压缩类
 * @auther: liuzhiyong
 * @date: 2018/8/28
 * @time: 下午 3:21
 */
public class ZipUtil {

    /**
     * @methodName: zipFiles
     * @param: [srcfiles 原文件集合, zipfile 压缩后的文件]
     * @describe: 将多个文件压缩成一个文件
     * @auther: liuzhiyong
     * @date: 2018/7/31
     * @time: 上午 10:15
     */
    public static boolean zipFiles(List<File> srcfiles, File zipfile) throws Exception{
        boolean sign = false;
        byte[] buf = new byte[1024];
        try {
            FileInputStream inputStream;
            ZipOutputStream out = new ZipOutputStream(new FileOutputStream(zipfile));
            for (int i = 0; i < srcfiles.size(); i++) {
                inputStream = new FileInputStream(srcfiles.get(i));
                out.putNextEntry(new ZipEntry(srcfiles.get(i).getName()));
                int len;
                while ((len = inputStream.read(buf)) > 0) {
                    out.write(buf, 0, len);
                }
                out.closeEntry();
                inputStream.close();
            }
            out.close();
            sign = true;
        } catch (Exception e) {
            System.out.println("文件压缩失败: " + e.getMessage());
        }
        return sign;
    }
}

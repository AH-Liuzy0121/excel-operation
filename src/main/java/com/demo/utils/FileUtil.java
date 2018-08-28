package com.demo.utils;

import java.io.File;

/**
 * @className: FileUtil
 * @package: com.demo.utils
 * @describe: 文件操作类
 * @auther: liuzhiyong
 * @date: 2018/8/28
 * @time: 下午 3:18
 */
public class FileUtil {

    /**
     * @methodName: deleteFiles
     * @param: [filePath 文件目录路径]
     * @describe: 删除该目录filePath下的所有文件
     * @auther: liuzhiyong
     * @date: 2018/8/28
     * @time: 下午 3:19
     */
    public static void deleteFiles(String filePath) {
        File file = new File(filePath);
        if (file.exists()) {
            File[] files = file.listFiles();
            for (int i = 0; i < files.length; i++) {
                if (files[i].isFile()) {
                    files[i].delete();
                }
            }
        }
    }

    /**
     * @methodName: deleteFile
     * @param: [filePath 文件目录路径, fileName 文件名称]
     * @describe: 删除单个文件
     * @auther: liuzhiyong
     * @date: 2018/8/28
     * @time: 下午 3:19
     */
    public static void deleteFile(String filePath, String fileName) {
        File file = new File(filePath);
        if (file.exists()) {
            File[] files = file.listFiles();
            for (int i = 0; i < files.length; i++) {
                if (files[i].isFile()) {
                    if (files[i].getName().equals(fileName)) {
                        files[i].delete();
                        return;
                    }
                }
            }
        }
    }
}

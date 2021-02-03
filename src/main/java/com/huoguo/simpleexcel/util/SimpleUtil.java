package com.huoguo.simpleexcel.util;

import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

/**
 * MultipartFile转换File对象
 *
 * @author Lizhenghuang
 */
public final class SimpleUtil {

    private static final int ZERO = 0, THOUSAND = 8192, MINUS = -1;

    /**
     * MultipartFile转换File对象
     *
     * @param file 文件对象
     * @return File
     */
    public static File multipartFileToFile(MultipartFile file) {
        File toFile = null;
        if (file.equals("") || file.getSize() <= 0) {
            file = null;
        } else {
            InputStream ins = null;
            try {
                ins = file.getInputStream();
                toFile = new File(file.getOriginalFilename());
                inputStreamToFile(ins, toFile);
                ins.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return toFile;
    }

    /**
     * Input输入流转换File对象
     *
     * @param ins  输入流
     * @param file 文件对象
     */
    private static void inputStreamToFile(InputStream ins, File file) {
        try {
            OutputStream os = new FileOutputStream(file);
            int bytesRead = ZERO;
            byte[] buffer = new byte[THOUSAND];
            while ((bytesRead = ins.read(buffer, ZERO, THOUSAND)) != MINUS) {
                os.write(buffer, ZERO, bytesRead);
            }
            os.close();
            ins.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 字符串过滤,中文符号转英文符号并去空格
     *
     * @param str 输入字符串
     * @return 字符串
     */
    public static String strFilter(String str) {
        return str.replace("！", "!")
                .replace("，", ",")
                .replace("。", "")
                .replace("；", "")
                .replace("：", ":")
                .trim();
    }

    /**
     * 字符串分割
     *
     * @param str       输入字符串
     * @param separator 分隔符
     * @return 字符串数组
     */
    public static String[] strSplit(String str, String separator) {
        return !StringUtils.isEmpty(separator) ? str.split(separator) : null;
    }
}

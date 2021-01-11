package com.huoguo.simpleexcel.util;

import org.springframework.web.multipart.MultipartFile;

import java.io.*;

/**
 * MultipartFile转换File对象
 * @author Lizhenghuang
 */
public final class MultipartFileToFile {

    private static final int ZERO = 0, THOUSAND = 8192, MINUS = -1;

    /**
     * MultipartFile转换File对象
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
     * @param ins 输入流
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
}

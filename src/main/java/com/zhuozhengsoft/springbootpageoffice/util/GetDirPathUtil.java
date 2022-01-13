package com.zhuozhengsoft.springbootpageoffice.util;

import org.springframework.util.ResourceUtils;

import java.io.FileNotFoundException;

/**
 * @Author: dong
 * @Date: 2020/11/2 10:39
 * @Version 1.0
 */
public class GetDirPathUtil {
    public static String getDirPath ()  {
        String path;
        try {
            path= ResourceUtils.getURL("classpath:").getPath();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new RuntimeException("没有找到项目根目录!");
        }
        return path;
    }
}

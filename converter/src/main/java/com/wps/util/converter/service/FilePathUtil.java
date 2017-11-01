package com.wps.util.converter.service;

import java.io.File;

/**
 * @author yang
 * @create_at 17-11-1
 **/
public class FilePathUtil {
    public static String getPathById(){
        String path=ClassLoader.getSystemResource("").getPath();
        path=path+"template//";
        return  path;
    }
}

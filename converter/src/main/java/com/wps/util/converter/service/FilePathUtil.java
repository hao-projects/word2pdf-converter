package com.wps.util.converter.service;


/**
 * @author yang
 * @create_at 17-11-1
 **/
public class FilePathUtil {
    public static String getPathById(){
        //String path=ClassLoader.getSystemResource("").getPath();
//        path=path+"template//";
        String path = FileService.class.getClassLoader().getResource("template").getPath();
        System.out.println(path);
        return  path;
    }
}

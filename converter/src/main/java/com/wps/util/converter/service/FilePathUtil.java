package com.wps.util.converter.service;


/**
 * @author yang
 * @create_at 17-11-1
 **/
public class FilePathUtil {
    public static String getPathById(){
        //String path=ClassLoader.getSystemResource("").getPath();
//        path=path+"static//";
        String path = FileService.class.getClassLoader().getResource("static").getPath();
        System.out.println(path);
        return  path;
    }
    // 判断一个字符是否是中文
    public static boolean isChinese(char c) {
        return c >= 0x4E00 &&  c <= 0x9FA5;// 根据字节码判断
    }
    // 判断一个字符串是否含有中文
    public static boolean isChinese(String str) {
        if (str == null) {return false;}
        for (char c : str.toCharArray()) {
            if (isChinese(c)) {return true;}// 有一个中文字符就返回
        }
        return false;
    }
}

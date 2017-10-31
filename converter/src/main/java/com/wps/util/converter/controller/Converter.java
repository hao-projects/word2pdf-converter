package com.wps.util.converter.controller;
import com.wps.util.converter.service.FileService;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Controller

public class Converter {
    private static Pattern pattern=Pattern.compile("(.*)->\\s(\\S*)(.*)(\\n*)(.*)");
    @RequestMapping(value = "/getpdf/{formType}")
    public void download(@RequestBody Map<String,Object> formMap, @PathVariable("formType") int formType, HttpServletResponse response)throws Exception{
       // String path =converter(request);
        System.out.println(formMap.toString());
        FileService fileService=new FileService();
        String path=fileService.file(formMap,formType);
        String result=executeCommand(path,"/tmp/");

        if(result!=null&&!result.contains("Error")){
            System.out.println("command result:"+result.trim());
            Matcher matcher=pattern.matcher(result.trim());
            System.out.println(matcher.matches());
            System.out.println("group0:"+matcher.group(0));

            result=matcher.group(2);
            System.out.println("group1"+result);

        }else{
            throw new Exception("can't not convert to pdf,error format");
        }
        File file = new File(result);
        Random random=new Random();
        int a=random.nextInt(50);
        String fileName=a+"download.pdf";
        //检查applyid是否是下载者的
        //validate.isPermission(SecurityUtils.getSubject(),fileService.getFileById(file_id).getApplyId());
        //String agent = request.getHeader("User-Agent");
        //String fileName = fileData.getFileName();
       // String fileName = "test";
        response.setHeader("content-type", "application/pdf");
        response.setContentType("application/pdf");
        response.setHeader("Content-Disposition", "inline;filename=\"" + fileName + "\"");
        OutputStream os=response.getOutputStream();
        FileInputStream fo=new FileInputStream(file);
        try{
            byte[] bis=readStream(fo);
            os.write(bis);
            os.flush();
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }finally {
            fo.close();
            os.close();
        }
    }
    public String converter(MultipartHttpServletRequest request)throws Exception{
        MultipartFile file = request.getFile("file");
        Random random=new Random();
        int a=random.nextInt(50);
        String filename=a+".docx";
        FileOutputStream fos=new FileOutputStream(new File("/tmp/"+filename));
        FileInputStream fs = (FileInputStream) file.getInputStream();

        byte[] buffer = new byte[1024];
        int len = 0;
        while ((len = fs.read(buffer)) != -1) {
            fos.write(buffer, 0, len
        );
        }
        fos.close();
        fs.close();
        String result=executeCommand("/tmp/"+filename,"/tmp/");

        if(result!=null&&!result.contains("Error")){
            System.out.println("command result:"+result.trim());
            Matcher matcher=pattern.matcher(result.trim());
            System.out.println(matcher.matches());
            System.out.println("group0:"+matcher.group(0));

            result=matcher.group(2);
            System.out.println("group1"+result);

        }else{
            throw new Exception("can't not convert to pdf,error format");
        }
        return result;

    }
    public static String executeCommand(String input_path,String output_folder_path){
       return execute("soffice --invisible --convert-to pdf --outdir " +output_folder_path+"  "+input_path);
    }

    public static String execute(String command) {
        String[] cmd = {"/bin/bash"};
        Runtime rt = Runtime.getRuntime();
        Process proc = null;
        try {
            proc = rt.exec(cmd);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 打开流
        OutputStream os = proc.getOutputStream();
        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(os));
        String result=null;
        try {
            bw.write(command);

            bw.flush();
            bw.close();

            /** 真奇怪，把控制台的输出打印一遍之后竟然能正常终止了~ */
            result=readConsole(proc);

            /** waitFor() 的作用在于 java 程序是否等待 Terminal 执行脚本完毕~ */
            proc.waitFor();
        } catch (Exception e) {
            e.printStackTrace();
        }
        int retCode = proc.exitValue();
//       System.err.println("unix script retCode = " + retCode);
        if (retCode != 0) {
            readConsole(proc);
            System.err.println("UnixScriptUil.execute 出错了!!");
        }
        return result;
    }

    /**
     * 读取控制命令的输出结果
     * 原文链接：http://lavasoft.blog.51cto.com/62575/15599
     * @return 控制命令的输出结果
     * @throws IOException
     */
    public static String readConsole(Process process) {
        StringBuffer cmdOut = new StringBuffer();
        InputStream fis = process.getInputStream();
        BufferedReader br = new BufferedReader(new InputStreamReader(fis));

        String line = null;
        try {
            while ((line = br.readLine()) != null) {
                cmdOut.append(line).append(System.getProperty("line.separator"));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("执行系统命令后的控制台输出为：\n" + cmdOut.toString());
        return cmdOut.toString();
    }
    public static byte[] readStream(InputStream inStream) {
        ByteArrayOutputStream bops = new ByteArrayOutputStream();
        int data = -1;
        try {
            while((data = inStream.read()) != -1){
                bops.write(data);
            }
            return bops.toByteArray();
        }catch(Exception e){
            return null;
        }
    }
}

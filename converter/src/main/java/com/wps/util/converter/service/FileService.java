package com.wps.util.converter.service;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBaseStylesOverride;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author yang
 * @create_at 17-10-22
 **/
public class FileService {

    public String file(Map<String,Object> replacetor,int form_type) throws Exception{
        String OUTPUT_PATH=FilePathUtil.getPathById()+"/test.docx";
        String INPUT_PATH=FilePathUtil.getPathById()+"/form"+form_type+".docx";
        System.out.println(form_type);
        try {

            InputStream is = new FileInputStream(INPUT_PATH);

            XWPFDocument doc=new XWPFDocument(is);
            if(form_type==2){
                form2Sublist(doc,(List<Object>)replacetor.get("subList"),4,0);

            }
            if(form_type==5){
                form2Sublist(doc,(List<Object>)replacetor.get("subList"),5,0);
            }
            if(form_type==7||form_type==6){
                generateCylindersOrPipes(doc,(List<Object>)replacetor.get("subList"),replacetor);
            }else{
                this.replaceInPara(doc,replacetor);
                this.replaceInTable(doc,replacetor);
            }

//            if(replacetor.get("subList")!=null)
//            {
//                Object a=replacetor.get("subList");
//                List<Object> b=(List<Object>) a;
//                for(Object c:b){
//                    Map<String,String>d=(Map<String,String>)c;
//
//                }


//            }
//			List<XWPFTable> tables=doc.getTables();
//			XWPFTable table=tables.get(0);
//			XWPFTableRow header = table.getRow(1);
//			table.addNewRowBetween(0,1);
//			table.addRow(header,2);
//			header.getCell(1).removeParagraph(0);
//			XWPFRun run= header.getCell(1).addParagraph().createRun();
//			header.getCell(3).removeParagraph(0);
//			XWPFRun run2= header.getCell(3).addParagraph().createRun();
//			run2.setText("test");
//			run2.setFontSize(10);
//			run.setText("test");
//			run.setFontSize(10);



            OutputStream os = new FileOutputStream(OUTPUT_PATH);
            doc.write(os);
            os.close();
            is.close();
        }
        catch (Exception e){
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
        return OUTPUT_PATH;
    }
    private void form2Sublist(XWPFDocument doc,List<Object> list,int list_pos,int table_index){
        XWPFTable table=doc.getTables().get(table_index);

        int i=0;
        for(Object str:list){
            Map<String,Object> map=(Map<String,Object>)str;
            map.put("iid",i+1);
            CTRow ctRow=CTRow.Factory.newInstance();
            ctRow.set(table.getRow(list_pos).getCtRow());
            XWPFTableRow row=new XWPFTableRow(ctRow,table);
            for(XWPFTableCell cell:row.getTableCells())
            {
                for(XWPFParagraph para:cell.getParagraphs()){
                    replaceInPara(para,map);
                }

            }
            i++;
            table.addRow(row,list_pos+i);
            if(i==9){
                break;
            }
        }
        table.removeRow(list_pos);
    }

    private void generateCylindersOrPipes(XWPFDocument doc, List<Object> list, Map<String, Object> replacetor)throws  Exception{


        int page_size=list.size()>list.size()/10*10?(list.size()/10+1):list.size()/10;

        CTBody body=doc.getDocument().getBody();
        replacetor.put("pageSize",page_size);
        replacetor.put("pageNum",1);
        form2Sublist(doc,list.subList(0,list.size()),1,0);
        replaceInPara(doc,replacetor);
        replaceInTable(doc,replacetor);
        for(int i=1;i<page_size;i++){
            XWPFDocument doc2=new XWPFDocument(doc.getPackage());
            form2Sublist(doc2,list.subList(10*i,list.size()),1,0);
            CTBody body2=doc2.getDocument().getBody();
            replacetor.put("pageNum",i+1);
            replaceInPara(doc2,replacetor);
            replaceInTable(doc2,replacetor);
            appendBody(body,body2);
        }

//        for(int j=0;j<page_size;j++){
//            form2Sublist(doc,list.subList(j*10,list.size()),1,j);
//        }
    }
    private static void appendBody(CTBody src, CTBody append) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);
        String srcString = src.xmlText();
        String prefix = srcString.substring(0,srcString.indexOf(">")+1);
        String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
        String sufix = srcString.substring( srcString.lastIndexOf("<") );
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
        src.set(makeBody);
    }
        /**
         * 替换段落里面的变量
         *
         * @param doc    要替换的文档
         * @param params 参数
         */
    private void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params);
        }
    }

    private void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
        List<XWPFRun> runs;
        Matcher matcher;
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                System.out.println(runText);
                matcher = this.matcher(runText);
                if (matcher.find()) {
                    while ((matcher = this.matcher(runText)).find()) {
                        Object ob=params.get(matcher.group(1));
                        if(ob!=null&&ob.toString()!=""){
                            runText = matcher.replaceFirst(String.valueOf(ob));
                        }
                        else{
                            runText= matcher.replaceFirst("-");
                        }

                        //System.out.println("find:" + matcher.group(1) + "end" + String.valueOf(params.get("second")));
                    }
                    //直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                    //所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                    para.removeRun(i);
                    if((FilePathUtil.isChinese(runText)&&runText.length()>8)||runText.length()>=17){
                        para.insertNewRun(i).setFontSize(7);
                        para.getRuns().get(i).setText(runText);
                    }
                    else{
                        para.insertNewRun(i).setText(runText);
                    }
                }
            }
        }
    }

    /**
     * 替换表格里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    private void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        this.replaceInPara(para, params);
                    }
                }
            }
        }
    }

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private static Pattern pattern=Pattern.compile("\\$\\{([^}]*)}", Pattern.CASE_INSENSITIVE);
    private Matcher matcher(String str) {
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

}

package com.wps.util.converter;

import com.sun.javafx.scene.shape.PathUtils;
import com.wps.util.converter.service.FilePathUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlOptions;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The type Converter application tests.
 */
//@RunWith(SpringRunner.class)
@SpringBootTest
public class ConverterApplicationTests {
	/**
	 * Context loads.
	 *
	 * @throws Exception the exception
	 */

//	@Test
//	public void word2pdf() throws Exception{
//		OpenOfficePDFConverter converter=new OpenOfficePDFConverter();
//		File input=new File("/home/ctt/Downloads/form1.docx");
//		File output=new File("/home/ctt/Downloads/formtest.pdf");
//		converter.convert2PDF("/home/ctt/Downloads/form1.docx","/home/ctt/Downloads/formtest.pdf");
//	}
	@Test
	public void contextLoads() throws Exception{
//		try {
//
//			//InputStream is = new FileInputStream("/home/ctt/Downloads/test2.docx");
////			InputStream is2 = new FileInputStream("/home/ctt/yang-workspace/web/word2pdf-converter/converter/src/main/resources/static/form1.docx");
//			OPCPackage pkg=OPCPackage.open(new FileInputStream("/home/ctt/yang-workspace/web/word2pdf-converter/converter/src/main/resources/static/form1.docx"));
//			Map<String,Object> replacetor=new HashMap<String,Object>();
//			replacetor.put("registKind","kkkkkk");
//			replacetor.put("eqCode","dianshu");
//			replacetor.put("manuComName","dierg");
//			replacetor.put("manufactureDate","hhhh");
//			Map<String,Object> replacetor2=new HashMap<String,Object>();
//			replacetor2.put("eqCode","好型急死了都放假了时间阿斯顿发生");
//			replacetor2.put("manuComName","dierg2");
//			replacetor2.put("manufactureDate","hhhh2");
//			List<Object> list=new ArrayList<Object>();
//			list.add(replacetor);
//			list.add(replacetor2);
//			list.add(replacetor);
//			XWPFDocument doc=new XWPFDocument(pkg);
//			XWPFDocument doc2=new XWPFDocument(doc.getPackage());
//			CTBody body=doc.getDocument().getBody();
//			CTBody body1=doc2.getDocument().getBody();
//			//appendBody(body,body1);
//			//appendBody(body,body1);
////			XWPFDocument doc2=new XWPFDocument(is2);
////			CTBody body2=doc2.getDocument().getBody();
//			List<XWPFTable> tables=doc.getTables();
//			XWPFTable table=tables.get(0);
//			int i=4;
//			List<XWPFTableRow> header=new ArrayList<XWPFTableRow>();
////
////			for(Object str:list){
////				System.out.println(i);
////				Map<String,Object> map=(Map<String,Object>)str;
////				//System.out.println(j);
////				CTRow ctRow=CTRow.Factory.newInstance();
////				ctRow.set(table.getRow(4).getCtRow());
////				XWPFTableRow row=new XWPFTableRow(ctRow,table);
////				for(XWPFTableCell cell:row.getTableCells())
////				{
////					for(XWPFParagraph para:cell.getParagraphs()){
////						replaceInPara(para,map);
////					}
////
////				}
////				table.addRow(row,5);
////
////			}
////			table.removeRow(4);
//
//
//			//table.addNewRowBetween(0,1);
//
//
//			// header.getCell(1).removeParagraph(0);
////			XWPFRun run= header.getCell(1).addParagraph().createRun();
////			header.getCell(3).removeParagraph(0);
////			XWPFRun run2= header.getCell(3).addParagraph().createRun();
////			run2.setText("test");
////			run2.setFontSize(10);
////			run.setText("test");
////			run.setFontSize(10);
////
//			this.replaceInPara(doc,replacetor);
//			this.replaceInTable(doc,replacetor,0);
//			XWPFDocument doc3=new XWPFDocument(doc.getPackage());
//			this.replaceInPara(doc2,replacetor);
//			this.replaceInTable(doc2,replacetor,0);
//			this.replaceInPara(doc3,replacetor);
//			this.replaceInTable(doc3,replacetor,0);
//			appendBody(body,body1);
//			appendBody(body,doc3.getDocument().getBody());
//			OutputStream os = new FileOutputStream("/home/ctt/Downloads/test2.docx");
//			doc.write(os);
//			os.close();
//			doc2.close();
//			doc.close();
//			doc3.close();
//			pkg.close();
//		}
//		catch (Exception e){
//			System.out.println(e.getMessage());
//			e.printStackTrace();
//		}
		System.out.println(FilePathUtil.getPathById());
	}

	private static void appendBody(CTBody src, CTBody append) throws Exception {
		try{XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = append.xmlText(optionsOuter);
		String srcString = src.xmlText();
		String prefix = srcString.substring(0,srcString.indexOf(">")+1);
		String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
		String sufix = srcString.substring( srcString.lastIndexOf("<") );
		String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
		CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
		src.set(makeBody);}
		catch (Exception e){
			src=append;
		}
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
				System.out.println("-----");
				matcher = this.matcher(runText);
				if (matcher.find()) {
					while ((matcher = this.matcher(runText)).find()) {
						runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
						System.out.println("find:" + matcher.group(1) + "end" + String.valueOf(params.get("second")));
					}
					//直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
					//所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
					para.removeRun(i);
					if((FilePathUtil.isChinese(runText)&&runText.length()>8)||runText.length()>=17){
						para.insertNewRun(i).setFontSize(7);
						para.getRuns().get(i).setText(runText);
					}
					else
						para.insertNewRun(i).setText(runText);
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
	 * replace in specific table with index
	 * @param doc document you want to change
	 * @param params	replace text map list
	 * @param tableIndex	table Index
	 */
	private void replaceInTable(XWPFDocument doc, Map<String, Object> params,int tableIndex) {
		XWPFTable table;
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;
		List<XWPFParagraph> paras;

			table = doc.getTables().get(tableIndex);
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

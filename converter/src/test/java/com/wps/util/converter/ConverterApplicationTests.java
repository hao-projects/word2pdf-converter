package com.wps.util.converter;

import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ConverterApplicationTests {
	@Test
	public void contextLoads() throws Exception{
		try {

			InputStream is = new FileInputStream("/home/ctt/Downloads/form2.docx");
			Map<String,Object> replacetor=new HashMap<String,Object>();
			replacetor.put("companyCode","dianshu");
			replacetor.put("manuComName","dierg");
			replacetor.put("manufactureDate","hhhh");
			Map<String,Object> replacetor2=new HashMap<String,Object>();
			replacetor2.put("companyCode","dianshu2");
			replacetor2.put("manuComName","dierg2");
			replacetor2.put("manufactureDate","hhhh2");
			List<Object> list=new ArrayList<Object>();
			list.add(replacetor);
			list.add(replacetor2);
			list.add(replacetor);
			XWPFDocument doc=new XWPFDocument(is);
			List<XWPFTable> tables=doc.getTables();
			XWPFTable table=tables.get(0);
			int i=4;
			List<XWPFTableRow> header=new ArrayList<XWPFTableRow>();

			for(Object str:list){
				System.out.println(i);
				Map<String,Object> map=(Map<String,Object>)str;
				//System.out.println(j);
				CTRow ctRow=CTRow.Factory.newInstance();
				ctRow.set(table.getRow(4).getCtRow());
				XWPFTableRow row=new XWPFTableRow(ctRow,table);
				for(XWPFTableCell cell:row.getTableCells())
				{
					for(XWPFParagraph para:cell.getParagraphs()){
						replaceInPara(para,map);
					}

				}
				table.addRow(row,5);

			}
			table.removeRow(4);


			//table.addNewRowBetween(0,1);


			// header.getCell(1).removeParagraph(0);
//			XWPFRun run= header.getCell(1).addParagraph().createRun();
//			header.getCell(3).removeParagraph(0);
//			XWPFRun run2= header.getCell(3).addParagraph().createRun();
//			run2.setText("test");
//			run2.setFontSize(10);
//			run.setText("test");
//			run.setFontSize(10);

		this.replaceInPara(doc,replacetor);
		this.replaceInTable(doc,replacetor);

			OutputStream os = new FileOutputStream("/home/ctt/Downloads/test2.docx");
			doc.write(os);
			os.close();
			is.close();
		}
		catch (Exception e){
			System.out.println(e.getMessage());
			e.printStackTrace();
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

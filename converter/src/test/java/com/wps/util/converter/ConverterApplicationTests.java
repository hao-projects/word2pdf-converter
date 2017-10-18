package com.wps.util.converter;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ConverterApplicationTests {

	@Test
	public void contextLoads() {
		String source="convert /tmp/39.docx -> /tmp//39.pdf using filter : writer_pdf_Export "
				+"ddddd";
		Matcher matcher= Pattern.compile("(.*)->\\s(\\S*)(.*)(\\n*)(.*)").matcher(source);
		System.out.println(matcher.matches());
		System.out.println("group0:"+matcher.group(2));

		//String result=matcher.group(1);
		//System.out.println("group1"+result);
	}

}

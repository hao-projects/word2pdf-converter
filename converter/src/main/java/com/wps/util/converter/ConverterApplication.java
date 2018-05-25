package com.wps.util.converter;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.boot.web.support.SpringBootServletInitializer;

@SpringBootApplication
public class ConverterApplication extends SpringBootServletInitializer {
	@Override
	protected SpringApplicationBuilder configure(SpringApplicationBuilder application) {
		return application.sources(ConverterApplication.class);
	}

	public static void main(String[] args) throws Exception {
//		System.out.println(System.getProperty("user.dir"));
		SpringApplication.run(ConverterApplication.class, args);
	}


}

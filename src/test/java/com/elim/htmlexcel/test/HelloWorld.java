package com.elim.htmlexcel.test;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;

import org.testng.annotations.Test;

import com.elim.htmlexcel.core.ExcelGenerator;

public class HelloWorld {

	@Test
	public void test() throws IOException {
		ExcelGenerator generator = new ExcelGenerator();
		InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream("helloworld.html");
		System.out.println(inputStream);
		BufferedReader br = new BufferedReader(new InputStreamReader(inputStream));
		StringBuilder builder = new StringBuilder();
		String line = null;
		while ((line = br.readLine()) != null) {
			builder.append(line);
		}
		String html = builder.toString();
		try (OutputStream out = new FileOutputStream("D:/helloworld.xlsx");) {
			generator.generate(html, out);
		}
	}
	
}

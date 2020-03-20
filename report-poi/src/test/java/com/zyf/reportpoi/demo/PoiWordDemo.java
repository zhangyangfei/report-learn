package com.zyf.reportpoi.demo;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

//import com.deepoove.poi.XWPFTemplate;
//import com.deepoove.poi.config.Configure;
//import com.deepoove.poi.config.Configure.ConfigureBuilder;

public class PoiWordDemo {

	public static void main(String[] args) throws IOException {
		System.out.println("111");
		test();
	}

	public static void test() throws IOException {
		// 报表业务数据集合
		Map<String, Object> reportDataMap = new HashMap<>();
		// 绑定模板
//		ConfigureBuilder builder = Configure.newBuilder();
////		builder.addPlugin('!', new TableRenderPolicy());
//		XWPFTemplate template = XWPFTemplate.compile("E:/poiword.docx",builder.build());
//		// 模板绑定数据
//		template.render(reportDataMap);
//		// 生成文件
//		template.writeToFile("E:/poiwordreport.docx");
	}
}

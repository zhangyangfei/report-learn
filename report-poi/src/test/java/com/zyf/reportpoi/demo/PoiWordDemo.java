package com.zyf.reportpoi.demo;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.Configure.ConfigureBuilder;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.util.BytePictureUtils;

public class PoiWordDemo {

	public static void main(String[] args) throws IOException {
		System.out.println("开始");
		test();
		System.out.println("结束");
	}

	public static void test() throws IOException {
		// 报表业务数据集合
		Map<String, Object> reportDataMap = new HashMap<>();
		reportDataMap.put("TITLE", "文章标题");
		reportDataMap.put("TITLE1", "产品基本列表");
		List<List<String>> dataList = new ArrayList<>();
		dataList.add(Arrays.asList(new String[] { "id001", "商品001", "2020-01-01", "2020-08-01" }));
		dataList.add(Arrays.asList(new String[] { "id002", "商品002", "2020-01-02", "2020-08-02" }));
		dataList.add(Arrays.asList(new String[] { "id003", "商品003", "2020-01-03", "2020-08-03" }));
		dataList.add(Arrays.asList(new String[] { "id004", "商品004", "2020-01-04", "2020-08-04" }));
		dataList.add(Arrays.asList(new String[] { "id005", "商品005", "2020-01-04", "2020-08-04" }));
		dataList.add(Arrays.asList(new String[] { "id006", "商品006", "2020-01-04", "2020-08-04" }));
		reportDataMap.put("S002", dataList);
		
		reportDataMap.put("TITLE2", "新型冠状病毒");
		reportDataMap.put("PARAGRAPH1", "    2019新型冠状病毒（2019-nCoV），因2019年武汉病毒性肺炎病例而被发现，2020年1月12日被世界卫生组织命名。冠状病毒是一个大型病毒家族，已知可引起感冒以及中东呼吸综合征（MERS）和严重急性呼吸综合征（SARS）等较严重疾病。新型冠状病毒是以前从未在人体中发现的冠状病毒新毒株。\r\n" + 
				"    2019年12月以来，湖北省武汉市持续开展流感及相关疾病监测，发现多起病毒性肺炎病例，均诊断为病毒性肺炎/肺部感染。\r\n" + 
				"    人感染了冠状病毒后常见体征有呼吸道症状、发热、咳嗽、气促和呼吸困难等。在较严重病例中，感染可导致肺炎、严重急性呼吸综合征、肾衰竭，甚至死亡。目前对于新型冠状病毒所致疾病没有特异治疗方法。但许多症状是可以处理的，因此需根据患者临床情况进行治疗。此外，对感染者的辅助护理可能非常有效。 ");

		reportDataMap.put("TITLE3", "折线图");
		File picture = new File("./test.png");
		reportDataMap.put("PICTURE1", new PictureRenderData(700, 300, ".png", BytePictureUtils.getLocalByteArray(picture)));
		
		reportDataMap.put("TITLE4", "表格测试");
//		RowRenderData header = RowRenderData.build(new TextRenderData("ff0000", "姓名"), new TextRenderData("ff0000", "学历")); // header 可以为null
		Style style = new Style();
		style.setBold(true);
		style.setColor("ff0000");
		style.setFontSize(16);
		RowRenderData header = RowRenderData.build(
				new TextRenderData("姓名", style), 
				new TextRenderData("学历", style),
				new TextRenderData("年龄", style));
		RowRenderData row0 = RowRenderData.build("张三", "研究生", "24");
		RowRenderData row1 = RowRenderData.build("李四", "博士", "26");
		RowRenderData row2 = RowRenderData.build("王五", "小学生", "9");
		RowRenderData row3 = RowRenderData.build(new String[] { "赵六", "本科生", "20" });
		reportDataMap.put("S004", new MiniTableRenderData(header, Arrays.asList(row0, row1, row2, row3)));
		
		// 绑定模板
		ConfigureBuilder builder = Configure.newBuilder();
		builder.addPlugin('!', new TableRenderPolicy());// 自定义表格标签插件，模板中占位符以!作为标识
		XWPFTemplate template = XWPFTemplate.compile("./poiword.docx", builder.build());
		// 模板绑定数据
		template.render(reportDataMap);
		// 生成文件
		template.writeToFile("./poiwordreport.docx");
	}
}

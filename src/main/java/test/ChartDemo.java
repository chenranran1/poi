package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * @Description:图表测试Demo
 * @author chenran
 * @date 2020年7月1日
 */
public class ChartDemo {

	public static void main(String[] args) throws Exception {

		System.out.println("POI图表内容替换测试");
		// final String returnurl = "D:\\POI\\pie_result.docx"; // 结果文件
		// final String templateurl = "D:\\POI\\pie.docx"; // 模板文件（饼图）

//		final String returnurl = "D:\\POI\\bar_result.docx"; // 结果文件
//		final String templateurl = "D:\\POI\\bar.docx"; // 模板文件（条形图）
		
//		final String returnurl = "D:\\POI\\pillar_result.docx"; // 结果文件
//		final String templateurl = "D:\\POI\\pillar.docx"; // 模板文件（柱状图）
		
//		final String returnurl = "D:\\POI\\line_result.docx"; // 结果文件
//		final String templateurl = "D:\\POI\\line.docx"; // 模板文件（折线图）
		
//		final String returnurl = "D:\\POI\\barAndLine_result.docx"; // 结果文件
//		final String templateurl = "D:\\POI\\barAndLine.docx"; // 模板文件（柱状图和折线图组合）
		
		final String returnurl = "D:\\POI\\dynamicChart_result.docx"; // 结果文件
		final String templateurl = "D:\\POI\\dynamicChart.docx"; // 模板文件（动态表格）
		
		InputStream is = new FileInputStream(new File(templateurl));
		XWPFDocument doc = new XWPFDocument(is);
		doCharts(doc);

		// 保存结果文件
		try {
			File file = new File(returnurl);
			if (file.exists()) {
				file.delete();
			}
			FileOutputStream fos = new FileOutputStream(returnurl);
			doc.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void doCharts(XWPFDocument doc) throws FileNotFoundException {
		
		// 数据准备
		List<String> titleArr = new ArrayList<>();// 标题
		titleArr.add("费用项");
		titleArr.add("金额");

		List<String> fldNameArr = new ArrayList<>();// 字段名
		fldNameArr.add("item");
		fldNameArr.add("account");

		// 数据集合
		List<Map<String, String>> listItemsByType = new ArrayList<Map<String, String>>();

		// 第一行数据
		Map<String, String> base1 = new HashMap<String, String>();
		base1.put("item", "材料费用");
		base1.put("account", "500");

		// 第二行数据
		Map<String, String> base2 = new HashMap<String, String>();
		base2.put("item", "出差费用");
		base2.put("account", "300");

		// 第三行数据
		Map<String, String> base3 = new HashMap<String, String>();
		base3.put("item", "住宿费用");
		base3.put("account", "300");

		Map<String, String> base4 = new HashMap<String, String>();
		base4.put("item", "水电费用");
		base4.put("account", "200");

		listItemsByType.add(base1);
		listItemsByType.add(base2);
		listItemsByType.add(base3);
		listItemsByType.add(base4);

		// 获取word模板中的所有图表元素，用map存放
		// 为什么不用list保存：查看doc.getRelations()的源码可知，源码中使用了hashMap读取文档图表元素，
		// 对relations变量进行打印后发现，图表顺序和文档中的顺序不一致，也就是说relations的图表顺序不是文档中从上到下的顺序
		Map<String, POIXMLDocumentPart> chartsMap = new HashMap<String, POIXMLDocumentPart>();
		// 动态刷新图表
		List<POIXMLDocumentPart> relations = doc.getRelations();
		for (POIXMLDocumentPart poixmlDocumentPart : relations) {
			if (poixmlDocumentPart instanceof XWPFChart) { // 如果是图表元素
				String str = poixmlDocumentPart.toString();
				System.out.println("str：" + str);
				String key = str.replaceAll("Name: ", "").replaceAll(
						" - Content Type: application/vnd\\.openxmlformats-officedocument\\.drawingml\\.chart\\+xml",
						"").trim();
				System.out.println("key：" + key);

				chartsMap.put(key, poixmlDocumentPart);
			}
		}

		System.out.println("\n图表数量：" + chartsMap.size() + "\n");
		
		// 柱状图
//		 POIXMLDocumentPart pie = chartsMap.get("/word/charts/chart1.xml");
//		 POIWordChartUtils.replacePieCharts(pie, titleArr, fldNameArr,
//		 listItemsByType);
		
		// 条形图
//		 POIXMLDocumentPart poixmlDocumentPart0 =
//		 chartsMap.get("/word/charts/chart1.xml");
//		 POIWordChartUtils.replaceBarCharts(poixmlDocumentPart0, titleArr,
//		 fldNameArr, listItemsByType);

		// 柱状图
//		 POIXMLDocumentPart poixmlDocumentPart1 =
//		 chartsMap.get("/word/charts/chart1.xml");
//		 POIWordChartUtils.replaceBarCharts(poixmlDocumentPart1, titleArr,
//		 fldNameArr, listItemsByType);

		// 折线图
//		 POIXMLDocumentPart poixmlDocumentPart1 =
//		 chartsMap.get("/word/charts/chart1.xml");
//		 POIWordChartUtils.replaceLineCharts(poixmlDocumentPart1, titleArr,
//		 fldNameArr, listItemsByType);
		
		// 多系列测试（需要准备好对应系列数的模板）
		// doCharts3(chartsMap);

		// 柱状图和直线图组合
		// doCharts6(chartsMap);
		
		// 动态表格
		doDynamicChart(chartsMap);
	}

	public static void doCharts3(Map<String, POIXMLDocumentPart> chartsMap) {
		// 数据准备
		List<String> titleArr = new ArrayList<String>();// 标题
		titleArr.add("费用项");
		titleArr.add("老张");
		titleArr.add("老王");
		titleArr.add("老刘");

		List<String> fldNameArr = new ArrayList<String>();// 字段名
		fldNameArr.add("item");
		fldNameArr.add("zhang");
		fldNameArr.add("wang");
		fldNameArr.add("liu");

		// 数据集合
		List<Map<String, String>> listItemsByType = new ArrayList<Map<String, String>>();

		// 第一行数据
		Map<String, String> base1 = new HashMap<String, String>();
		base1.put("item", "1月");
		base1.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base1.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base1.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第二行数据
		Map<String, String> base2 = new HashMap<String, String>();
		base2.put("item", "2月");
		base2.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base2.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base2.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第三行数据
		Map<String, String> base3 = new HashMap<String, String>();
		base3.put("item", "3月");
		base3.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base3.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base3.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第三行数据
		Map<String, String> base4 = new HashMap<String, String>();
		base4.put("item", "3月");
		base4.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base4.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base4.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		
		listItemsByType.add(base1);
		listItemsByType.add(base2);
		listItemsByType.add(base3);
		listItemsByType.add(base4);

		POIXMLDocumentPart poixmlDocumentPart2 = chartsMap.get("/word/charts/chart1.xml");
		POIWordChartUtils.replaceBarCharts(poixmlDocumentPart2, titleArr, fldNameArr, listItemsByType);
	}
	
	public static void doCharts6(Map<String, POIXMLDocumentPart> chartsMap) {
		// 数据准备
		List<String> titleArr = new ArrayList<String>();// 标题
		titleArr.add("销售额");
		titleArr.add("老王");
		titleArr.add("老刘");

		List<String> fldNameArr = new ArrayList<String>();// 字段名
		fldNameArr.add("item");
		fldNameArr.add("wang");
		fldNameArr.add("liu");

		// 数据集合
		List<Map<String, String>> listItemsByType = new ArrayList<Map<String, String>>();

		// 第一行数据
		Map<String, String> base1 = new HashMap<String, String>();
		base1.put("item", "1月");
		base1.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base1.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第二行数据
		Map<String, String> base2 = new HashMap<String, String>();
		base2.put("item", "2月");
		base2.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base2.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第三行数据
		Map<String, String> base3 = new HashMap<String, String>();
		base3.put("item", "3月");
		base3.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base3.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第三行数据
		Map<String, String> base4 = new HashMap<String, String>();
		base4.put("item", "3月");
		base4.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base4.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		
		listItemsByType.add(base1);
		listItemsByType.add(base2);
		listItemsByType.add(base3);
		listItemsByType.add(base4);

		POIXMLDocumentPart poixmlDocumentPart2 = chartsMap.get("/word/charts/chart1.xml");
		POIWordChartUtils.replaceCombinationCharts(poixmlDocumentPart2, titleArr, fldNameArr, listItemsByType);
	}
	
	public static void doDynamicChart(Map<String, POIXMLDocumentPart> chartsMap) {
		// 数据准备
		List<String> titleArr = new ArrayList<String>();// 标题
		titleArr.add("费用项");
		titleArr.add("老张");
		titleArr.add("老王");
		titleArr.add("老刘");

		List<String> fldNameArr = new ArrayList<String>();// 字段名
		fldNameArr.add("item");
		fldNameArr.add("zhang");
		fldNameArr.add("wang");
		fldNameArr.add("liu");

		// 数据集合
		List<Map<String, String>> listItemsByType = new ArrayList<Map<String, String>>();

		// 第一行数据
		Map<String, String> base1 = new HashMap<String, String>();
		base1.put("item", "1月");
		base1.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base1.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base1.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第二行数据
		Map<String, String> base2 = new HashMap<String, String>();
		base2.put("item", "2月");
		base2.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base2.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base2.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第三行数据
		Map<String, String> base3 = new HashMap<String, String>();
		base3.put("item", "3月");
		base3.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base3.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base3.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");

		// 第三行数据
		Map<String, String> base4 = new HashMap<String, String>();
		base4.put("item", "3月");
		base4.put("zhang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base4.put("wang", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		base4.put("liu", (int) (1 + Math.random() * (100 - 1 + 1)) + "");
		
		listItemsByType.add(base1);
		listItemsByType.add(base2);
		listItemsByType.add(base3);
		listItemsByType.add(base4);

		POIXMLDocumentPart poixmlDocumentPart2 = chartsMap.get("/word/charts/chart1.xml");
		POIWordDynamicChartUtils.replaceBarCharts(poixmlDocumentPart2, titleArr, fldNameArr, listItemsByType);
	}

}

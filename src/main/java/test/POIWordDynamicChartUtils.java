package test;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSchemeColor;

/**
 * @Description:动态表格工具类
 * @author chenran
 * @date 2020年7月1日
 */
public class POIWordDynamicChartUtils {

	/**
	 * 调用替换柱状图数据-可以实现动态列 其它的折线图、饼图大同小异
	 */
	public static void replaceBarCharts(POIXMLDocumentPart poixmlDocumentPart, List<String> titleArr,
			List<String> fldNameArr, List<Map<String, String>> listItemsByType) {

		// 很重要的参数，图表系列的数量，由这里的数据决定
		int culomnNum = titleArr.size() - 1;

		XWPFChart chart = (XWPFChart) poixmlDocumentPart;
		chart.getCTChart();

		// 根据属性第一列名称切换数据类型
		CTChart ctChart = chart.getCTChart();
		CTPlotArea plotArea = ctChart.getPlotArea();

		// 柱状图用plotArea.getBarChartArray，
		// 折线图用plotArea.getLineChartArray，
		// 饼图用plotArea.getPieChartArray.....
		// 还有很多其它的
		CTBarChart barChart = plotArea.getBarChartArray(0);

		// 清除图表的样式，由代码自己设置
		barChart.getSerList().clear();

		// 刷新内置excel数据
		refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
		// 刷新页面显示数据-也就是刷新图表的数据源范围
		refreshBarStrGraphContent(barChart, listItemsByType, fldNameArr, 1, culomnNum, titleArr);
	}

	/**
	 * 调用替换柱状图数据-可以实现动态列 其它的折线图、饼图大同小异
	 */
	public static void replaceLineCharts(POIXMLDocumentPart poixmlDocumentPart, List<String> titleArr,
			List<String> fldNameArr, List<Map<String, String>> listItemsByType) {

		// 很重要的参数，图表系列的数量，由这里的数据决定
		int culomnNum = titleArr.size() - 1;

		XWPFChart chart = (XWPFChart) poixmlDocumentPart;
		chart.getCTChart();

		// 根据属性第一列名称切换数据类型
		CTChart ctChart = chart.getCTChart();
		CTPlotArea plotArea = ctChart.getPlotArea();

		// 柱状图用plotArea.getBarChartArray，
		// 折线图用plotArea.getLineChartArray，
		// 饼图用plotArea.getPieChartArray.....
		// 还有很多其它的
		CTLineChart ctLineChart = plotArea.getLineChartArray(0);

		// 清除图表的样式，由代码自己设置
		ctLineChart.getSerList().clear();

		// 刷新内置excel数据
		refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
		// 刷新页面显示数据-也就是刷新图表的数据源范围
		refreshLineStrGraphContent(ctLineChart, listItemsByType, fldNameArr, 1, culomnNum, titleArr);
	}

	/**
	 * 动态添加列-柱状图的
	 */
	public static boolean refreshBarStrGraphContent(CTBarChart barChart, List<Map<String, String>> dataList,
			List<String> fldNameArr, int position, int culomnNum, List<String> titleArr) {
		boolean result = true;
		// 更新数据区域
		for (int i = 0; i < culomnNum; i++) {
			CTBarSer ctBarSer = barChart.addNewSer();
			ctBarSer.addNewIdx().setVal(i);
			ctBarSer.addNewOrder().setVal(i);

			// 设置柱状图的系列名称
			// 设置标题 用以下这个方式，可以兼容office和wps（因为是动态添加，不可以直接get到，需要这样写）
			CTSerTx tx = ctBarSer.addNewTx();
			CTStrRef ctStrRef = tx.addNewStrRef();
			CTStrData ctStrData = ctStrRef.addNewStrCache();
			ctStrData.addNewPtCount().setVal(1);
			CTStrVal ctStrVal = ctStrData.addNewPt();
			ctStrVal.setIdx(0);
			ctStrVal.setV(titleArr.get(i + 1)); // 设置系列的名称

			// 设置柱状图系列的颜色，就是显示的柱子的颜色，不设置的话会默认都是黄色
			// 必须使用ACCENT_x系列的才行
			CTSchemeColor ctSchemeColor = ctBarSer.addNewSpPr().addNewSolidFill().addNewSchemeClr();
			ctSchemeColor.setVal(STSchemeColorValTools.get(i));

			CTAxDataSource cat = ctBarSer.addNewCat();
			CTNumDataSource val = ctBarSer.addNewVal();

			CTStrData strData = cat.addNewStrRef().addNewStrCache();
			CTNumData numData = val.addNewNumRef().addNewNumCache();
			strData.setPtArray((CTStrVal[]) null); // unset old axis text
			numData.setPtArray((CTNumVal[]) null); // unset old values

			// set model
			long idx = 0;
			for (int j = 0; j < dataList.size(); j++) {
				// 判断获取的值是否为空
				String value = "0";
				if (new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))) != null) {
					value = new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))).toString();
				}
				if (!"0".equals(value)) {
					CTNumVal numVal = numData.addNewPt();// 序列值
					numVal.setIdx(idx);
					numVal.setV(value);
				}
				CTStrVal sVal = strData.addNewPt();// 序列名称
				sVal.setIdx(idx);
				sVal.setV(dataList.get(j).get(fldNameArr.get(0)));
				idx++;
			}

			numData.addNewPtCount().setVal(idx);
			strData.addNewPtCount().setVal(idx);

		}
		return result;
	}

	/**
	 * 动态添加列-折线图的
	 */
	public static boolean refreshLineStrGraphContent(CTLineChart ctLineChart, List<Map<String, String>> dataList,
			List<String> fldNameArr, int position, int culomnNum, List<String> titleArr) {
		boolean result = true;
		// 更新数据区域
		for (int i = 0; i < culomnNum; i++) {
			CTLineSer ctLineSer = ctLineChart.addNewSer();
			ctLineSer.addNewIdx().setVal(i);
			ctLineSer.addNewOrder().setVal(i);

			// 设置柱状图的系列名称
			// 设置标题 用以下这个方式，可以兼容office和wps（因为是动态添加，不可以直接get到，需要这样写）
			CTSerTx tx = ctLineSer.addNewTx();
			CTStrRef ctStrRef = tx.addNewStrRef();
			CTStrData ctStrData = ctStrRef.addNewStrCache();
			ctStrData.addNewPtCount().setVal(1);
			CTStrVal ctStrVal = ctStrData.addNewPt();
			ctStrVal.setIdx(0);
			ctStrVal.setV(titleArr.get(i + 1)); // 设置系列的名称

			// 设置柱状图系列的颜色，就是显示的柱子的颜色，不设置的话会默认都是黄色
			// 必须使用ACCENT_x系列的才行
			CTSchemeColor ctSchemeColor = ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSchemeClr();
			ctSchemeColor.setVal(STSchemeColorValTools.get(i));

			CTAxDataSource cat = ctLineSer.addNewCat();
			CTNumDataSource val = ctLineSer.addNewVal();

			CTStrData strData = cat.addNewStrRef().addNewStrCache();
			CTNumData numData = val.addNewNumRef().addNewNumCache();
			strData.setPtArray((CTStrVal[]) null); // unset old axis text
			numData.setPtArray((CTNumVal[]) null); // unset old values

			// set model
			long idx = 0;
			for (int j = 0; j < dataList.size(); j++) {
				// 判断获取的值是否为空
				String value = "0";
				if (new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))) != null) {
					value = new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))).toString();
				}
				if (!"0".equals(value)) {
					CTNumVal numVal = numData.addNewPt();// 序列值
					numVal.setIdx(idx);
					numVal.setV(value);
				}
				CTStrVal sVal = strData.addNewPt();// 序列名称
				sVal.setIdx(idx);
				sVal.setV(dataList.get(j).get(fldNameArr.get(0)));
				idx++;
			}

			numData.addNewPtCount().setVal(idx);
			strData.addNewPtCount().setVal(idx);

		}
		return result;
	}

	/**
	 * 刷新内置excel数据
	 *
	 * @param chart
	 * @param dataList
	 * @param fldNameArr
	 * @param titleArr
	 * @return
	 */
	public static boolean refreshExcel(XWPFChart chart, List<Map<String, String>> dataList, List<String> fldNameArr,
			List<String> titleArr) {
		boolean result = true;
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		// 根据数据创建excel第一行标题行
		for (int i = 0; i < titleArr.size(); i++) {
			if (sheet.getRow(0) == null) {
				sheet.createRow(0).createCell(i).setCellValue(titleArr.get(i) == null ? "" : titleArr.get(i));
			} else {
				sheet.getRow(0).createCell(i).setCellValue(titleArr.get(i) == null ? "" : titleArr.get(i));
			}
		}

		// 遍历数据行
		for (int i = 0; i < dataList.size(); i++) {
			Map<String, String> baseFormMap = dataList.get(i);// 数据行
			// fldNameArr字段属性
			for (int j = 0; j < fldNameArr.size(); j++) {
				if (sheet.getRow(i + 1) == null) {
					if (j == 0) {
						try {
							sheet.createRow(i + 1).createCell(j)
									.setCellValue(baseFormMap.get(fldNameArr.get(j)) == null ? ""
											: baseFormMap.get(fldNameArr.get(j)));
						} catch (Exception e) {
							if (baseFormMap.get(fldNameArr.get(j)) == null) {
								sheet.createRow(i + 1).createCell(j).setCellValue("");
							} else {
								sheet.createRow(i + 1).createCell(j).setCellValue(baseFormMap.get(fldNameArr.get(j)));
							}
						}
					}
				} else {
					BigDecimal b = new BigDecimal(baseFormMap.get(fldNameArr.get(j)));
					double value = 0d;
					if (b != null) {
						value = b.doubleValue();
					}
					if (value == 0) {
						sheet.getRow(i + 1).createCell(j);
					} else {
						sheet.getRow(i + 1).createCell(j).setCellValue(b.doubleValue());
					}
				}
			}

		}
		// 更新嵌入的workbook
		POIXMLDocumentPart xlsPart = chart.getRelations().get(0);
		OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();

		try {
			wb.write(xlsOut);
			xlsOut.close();
		} catch (IOException e) {
			e.printStackTrace();
			result = false;
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
					result = false;
				}
			}
		}
		return result;
	}

}

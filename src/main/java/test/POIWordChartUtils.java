package test;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
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
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;

/**
 * @Description:POI处理Word图标相关内容工具类（包括四种主要类型图表：分别是条形图、柱状图、折线图和饼图）
 * @author chenran
 * @date 2020年7月1日
 */
public class POIWordChartUtils {

	 /**
     * 刷新内置excel数据
     * @param chart
     * @param dataList
     * @param fldNameArr
     * @param titleArr
     * @return
     */
    public static boolean refreshExcel(XWPFChart chart,
                                       List<Map<String, String>> dataList, List<String> fldNameArr, List<String> titleArr) {
        boolean result = true;
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        //根据数据创建excel第一行标题行
        for (int i = 0; i < titleArr.size(); i++) {
            if (sheet.getRow(0) == null) {
                sheet.createRow(0).createCell(i).setCellValue(titleArr.get(i) == null ? "" : titleArr.get(i));
            } else {
                sheet.getRow(0).createCell(i).setCellValue(titleArr.get(i) == null ? "" : titleArr.get(i));
            }
        }
 
        //遍历数据行
        for (int i = 0; i < dataList.size(); i++) {
            Map<String, String> baseFormMap = dataList.get(i);//数据行
            //fldNameArr字段属性
            for (int j = 0; j < fldNameArr.size(); j++) {
                if (sheet.getRow(i + 1) == null) {
                    if (j == 0) {
                        try {
                            sheet.createRow(i + 1).createCell(j).setCellValue(baseFormMap.get(fldNameArr.get(j)) == null ? "" : baseFormMap.get(fldNameArr.get(j)));
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

    /**
     * 刷新饼图数据
     * @param typeChart
     * @param serList
     * @param dataList
     * @param fldNameArr
     * @param position
     * @return
     */
	public static boolean refreshPieStrGraphContent(Object typeChart, List<?> serList,
			List<Map<String, String>> dataList, List<String> fldNameArr, int position) {

		boolean result = true;
		// 更新数据区域
		for (int i = 0; i < serList.size(); i++) {
			// CTSerTx tx=null;
			CTAxDataSource cat = null;
			CTNumDataSource val = null;
			CTPieSer ser = ((CTPieChart) typeChart).getSerArray(i);

			// tx= ser.getTx();
			// Category Axis Data
			cat = ser.getCat();
			// 获取图表的值
			val = ser.getVal();
			// strData.set
			CTStrData strData = cat.getStrRef().getStrCache();
			CTNumData numData = val.getNumRef().getNumCache();
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
			numData.getPtCount().setVal(idx);
			strData.getPtCount().setVal(idx);

			// 赋值横坐标数据区域
			String axisDataRange = new CellRangeAddress(1, dataList.size(), 0, 0).formatAsString("Sheet1", true);
			cat.getStrRef().setF(axisDataRange);

			// 数据区域
			String numDataRange = new CellRangeAddress(1, dataList.size(), i + position, i + position)
					.formatAsString("Sheet1", true);
			val.getNumRef().setF(numDataRange);
		}
		return result;
	}

	/**
     * 调用替换饼图数据
     */
    public static void replacePieCharts(POIXMLDocumentPart poixmlDocumentPart,
                                        List<String> titleArr, List<String> fldNameArr, List<Map<String, String>> listItemsByType) {
        XWPFChart chart = (XWPFChart) poixmlDocumentPart;
 
        //根据属性第一列名称切换数据类型
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
 
        CTPieChart pieChart = plotArea.getPieChartArray(0);
        List<CTPieSer> pieSerList = pieChart.getSerList();  // 获取饼图单位
 
        //刷新内置excel数据
        refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
        //刷新页面显示数据
        refreshPieStrGraphContent(pieChart, pieSerList, listItemsByType, fldNameArr, 1);
 
    }
	
	 /**
     * 刷新柱状图数据方法
     *
     * @param typeChart
     * @param serList
     * @param dataList
     * @param fldNameArr
     * @param position
     * @return
     */
    public  static boolean refreshBarStrGraphContent(Object typeChart,
                                                    List<?> serList, List<Map<String, String>> dataList, List<String> fldNameArr, int position) {
        boolean result = true;
        //更新数据区域
        for (int i = 0; i < serList.size(); i++) {
//            CTSerTx tx=null;
            CTAxDataSource cat = null;
            CTNumDataSource val = null;
            CTBarSer ser = ((CTBarChart) typeChart).getSerArray(i);
//            tx= ser.getTx();
 
            // Category Axis Data
            cat = ser.getCat();
            // 获取图表的值
            val = ser.getVal();
            // strData.set
            CTStrData strData = cat.getStrRef().getStrCache();
            CTNumData numData = val.getNumRef().getNumCache();
            strData.setPtArray((CTStrVal[]) null); // unset old axis text
            numData.setPtArray((CTNumVal[]) null); // unset old values
 
            // set model
            long idx = 0;
            for (int j = 0; j < dataList.size(); j++) {
                //判断获取的值是否为空
                String value = "0";
                if (new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))) != null) {
                    value = new BigDecimal(dataList.get(j).get(fldNameArr.get(i + position))).toString();
                }
                if (!"0".equals(value)) {
                    CTNumVal numVal = numData.addNewPt();//序列值
                    numVal.setIdx(idx);
                    numVal.setV(value);
                }
                CTStrVal sVal = strData.addNewPt();//序列名称
                sVal.setIdx(idx);
                sVal.setV(dataList.get(j).get(fldNameArr.get(0)));
                idx++;
            }
            numData.getPtCount().setVal(idx);
            strData.getPtCount().setVal(idx);
 
 
            //赋值横坐标数据区域
            String axisDataRange = new CellRangeAddress(1, dataList.size(), 0, 0)
                    .formatAsString("Sheet1", true);
            cat.getStrRef().setF(axisDataRange);
 
            //数据区域
            String numDataRange = new CellRangeAddress(1, dataList.size(), i + position, i + position)
                    .formatAsString("Sheet1", true);
            val.getNumRef().setF(numDataRange);
 
        }
        return result;
    }

    /**
     * 调用替换柱状图数据
     */
    public static void replaceBarCharts(POIXMLDocumentPart poixmlDocumentPart,
                                        List<String> titleArr, List<String> fldNameArr, List<Map<String, String>> listItemsByType) {
        XWPFChart chart = (XWPFChart) poixmlDocumentPart;
 
        //根据属性第一列名称切换数据类型
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
 
        CTBarChart barChart = plotArea.getBarChartArray(0);
        List<CTBarSer> BarSerList = barChart.getSerList();  // 获取柱状图单位
 
        //刷新内置excel数据
        refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
        //刷新页面显示数据
        refreshBarStrGraphContent(barChart, BarSerList, listItemsByType, fldNameArr, 1);
    }
    
    /**
     * 刷新折线图数据方法
     * @param typeChart
     * @param serList
     * @param dataList
     * @param fldNameArr
     * @param position
     * @return
     */
    public static boolean refreshLineStrGraphContent(Object typeChart, List<?> serList,
			List<Map<String, String>> dataList, List<String> fldNameArr, int position) {

		boolean result = true;
		// 更新数据区域
		for (int i = 0; i < serList.size(); i++) {
			// CTSerTx tx=null;
			CTAxDataSource cat = null;
			CTNumDataSource val = null;
			CTLineSer ser = ((CTLineChart) typeChart).getSerArray(i);
			// tx= ser.getTx();
			// Category Axis Data
			cat = ser.getCat();
			// 获取图表的值
			val = ser.getVal();
			// strData.set
			CTStrData strData = cat.getStrRef().getStrCache();
			CTNumData numData = val.getNumRef().getNumCache();
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
			numData.getPtCount().setVal(idx);
			strData.getPtCount().setVal(idx);

			// 赋值横坐标数据区域
			String axisDataRange = new CellRangeAddress(1, dataList.size(), 0, 0).formatAsString("Sheet1", false);
			cat.getStrRef().setF(axisDataRange);

			// 数据区域
			String numDataRange = new CellRangeAddress(1, dataList.size(), i + position, i + position)
					.formatAsString("Sheet1", false);
			val.getNumRef().setF(numDataRange);

			// 设置系列生成方向

		}
		return result;
	}
	
    /**
     * 调用替换折线图数据
     */
    public static void replaceLineCharts(POIXMLDocumentPart poixmlDocumentPart,
                                         List<String> titleArr, List<String> fldNameArr, List<Map<String, String>> listItemsByType) {
        XWPFChart chart = (XWPFChart) poixmlDocumentPart;
 
        //根据属性第一列名称切换数据类型
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
 
        CTLineChart lineChart = plotArea.getLineChartArray(0);
        List<CTLineSer> lineSerList = lineChart.getSerList();   // 获取折线图单位
 
        //刷新内置excel数据
        refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
        //刷新页面显示数据
        refreshLineStrGraphContent(lineChart, lineSerList, listItemsByType, fldNameArr, 1);
 
    }
	
    /**
     * 调用替换柱状图、折线图组合数据
     */
    public static void replaceCombinationCharts(POIXMLDocumentPart poixmlDocumentPart,
                                                List<String> titleArr, List<String> fldNameArr, List<Map<String, String>> listItemsByType) {
        XWPFChart chart = (XWPFChart) poixmlDocumentPart;
        chart.getCTChart();
 
        //根据属性第一列名称切换数据类型
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
 
 
        CTBarChart barChart = plotArea.getBarChartArray(0);
        List<CTBarSer> barSerList = barChart.getSerList();  // 获取柱状图单位
        //刷新内置excel数据
        refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
        //刷新页面显示数据   数据中下标1开始的是柱状图数据，所以这个是1
        refreshBarStrGraphContent(barChart, barSerList, listItemsByType, fldNameArr, 1);
 
 
        CTLineChart lineChart = plotArea.getLineChartArray(0);
        List<CTLineSer> lineSerList = lineChart.getSerList();   // 获取折线图单位
        //刷新内置excel数据   有一个就可以了    有一个就可以了    有一个就可以了
        //refreshExcel(chart, listItemsByType, fldNameArr, titleArr);
        //刷新页面显示数据   数据中下标2开始的是折线图的数据，所以这个是2
        refreshLineStrGraphContent(lineChart, lineSerList, listItemsByType, fldNameArr, 2);
 
    }

}

package test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Description:POI操作Excel测试（主要是测试各种类型值的获取，其他操作与jxl类似）
 * @author chenran
 * @date 2020年7月1日
 */
public class ExcelDemo {

	public static void main(String[] args) throws Exception {
		
		// 读取表格
//		XSSFWorkbook workbook = new XSSFWorkbook(new File("D:\\POI\\excel.xlsx"));
//		XSSFSheet sheet = workbook.getSheetAt(0);
//		XSSFRow row = sheet.getRow(0);
//		String value1 = POIExcelUtils.getCellStringValue(row.getCell(0));
//		System.out.println(value1);
//		String value2 = POIExcelUtils.getCellStringValue(row.getCell(1));
//		System.out.println(value2);
//		Date date = POIExcelUtils.getCellDateValue(row.getCell(2));
//		System.out.println(date);
//		workbook.close();
		
		// 输出表格
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		sheet.createRow(0).createCell(0).setCellValue("输出表格内容");
		File file = new File("D:\\POI\\excel_result.xlsx");
		if (!file.exists()) {
			file.createNewFile();
		}
		workbook.write(new FileOutputStream(file));
		workbook.close();
	}
	
}

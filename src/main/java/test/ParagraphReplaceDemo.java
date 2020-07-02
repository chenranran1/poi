package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

/**
 * @Description:段落内容替换测试Demo（包括文本、插入图片和表格相关操作）
 * @author chenran
 * @date 2020年7月1日
 */
public class ParagraphReplaceDemo {

	public static void main(String[] args) throws Exception {
		System.out.println("POI替换段落内容测试");
		replaceParagraph();
	}

	/**
	 * 动态替换文本
	 * 
	 * @throws IOException
	 */
	public static void replaceParagraph() throws Exception {

		final String returnurl = "D:\\POI\\paragraph1_result.docx"; // 结果文件

		final String templateurl = "D:\\POI\\paragraph1.docx"; // 模板文件（必须有一个不为空的模板文件）

		InputStream is = new FileInputStream(new File(templateurl));
		XWPFDocument doc = new XWPFDocument(is);

		// 替换word中的段落模板数据
		doParagraphs(doc);

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

	public static void doParagraphs(XWPFDocument doc) throws Exception {

		// 文本数据
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("${text}", "我是被替换的文本内容");

		// 图片，如果是多个图片，就新建多个map
		Map<String, Object> imgMap = new HashMap<String, Object>();
		imgMap.put("width", 100);
		imgMap.put("height", 150);
		imgMap.put("type", "jpg");
		imgMap.put("path", "D:\\POI\\test.jpg");
		dataMap.put("${picture}", imgMap);

		POIWordTextUtils.changeText(doc, dataMap);
		
		/** ----------------------------定制化表格相关处理-------------------------- **/
		List<XWPFParagraph> paragraphList = doc.getParagraphs();
		if (paragraphList != null && paragraphList.size() > 0) {
			for (XWPFParagraph paragraph : paragraphList) {
				List<XWPFRun> runs = paragraph.getRuns();
				for (XWPFRun run : runs) {
					String text = run.getText(0);
					if (text != null && text.contains("${table}")) {

						// 动态表格
						run.setText("", 0);
						XmlCursor cursor = paragraph.getCTP().newCursor();
						XWPFTable tableOne = doc.insertNewTbl(cursor);// ---这个是关键

						// 设置表格宽度，第一行宽度就可以了，这个值的单位，目前我也还不清楚，还没来得及研究
						tableOne.setWidth(8500);

						// 表格第一行，对于每个列，必须使用createCell()，而不是getCell()，因为第一行嘛，肯定是属于创建的，没有create哪里来的get呢
						XWPFTableRow tableOneRowOne = tableOne.getRow(0);// 行
						POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(0), "微软雅黑", "9", 0, "left", "top",
								"#000000", "#FFFFFF", "20%", "序号");
						POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top",
								"#000000", "#FFFFFF", "40%", "公司名称(英文)");
						POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top",
								"#000000", "#FFFFFF", "40%", "公司名称(中文)");

						// 表格第二行
						XWPFTableRow tableOneRowTwo = tableOne.createRow();// 行
						POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo.getCell(0), "微软雅黑", "9", 0, "left", "top",
								"#000000", "#FFFFFF", "20%", "一行一列");
						POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo.getCell(1), "微软雅黑", "9", 0, "left", "top",
								"#000000", "#FFFFFF", "40%", "一行一列");
						POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo.getCell(2), "微软雅黑", "9", 0, "left", "top",
								"#000000", "#FFFFFF", "40%", "一行一列");

						// ....... 还可以动态添加表格
					}
				}
			}
		}
	}

}

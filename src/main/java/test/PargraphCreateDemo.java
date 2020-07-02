package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

/**
 * @Description:段落操作相关内容（包括创建文本、插入图片和表格相关操作）
 * @author chenran
 * @date 2020年7月1日
 */
public class PargraphCreateDemo {

	public static void main(String[] args) throws Exception {
		System.out.println("POI段落操作测试");
		createParagraph();
	}
	
	public static void createParagraph() throws Exception {
		
		final String returnurl = "D:\\POI\\paragraph_result.docx";  // 结果文件
		 
        final String templateurl = "D:\\POI\\paragraph.docx";  // 模板文件（模板文件不能为空文件）
 
        InputStream is = new FileInputStream(new File(templateurl));
        
        // 新创建文件，不需要模板
        XWPFDocument doc = new XWPFDocument();
        
        // 段落操作
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
	
	private static void doParagraphs(XWPFDocument doc) throws InvalidFormatException, FileNotFoundException, IOException {
		
		// 创建文本段落
		XWPFParagraph text_paragraph = doc.createParagraph();
		XWPFRun text_run = text_paragraph.createRun();
		text_run.setText("自己创建的文本段落");
		
		// 创建图片段落
		XWPFParagraph pic_paragraph = doc.createParagraph();
		XWPFRun pic_run = pic_paragraph.createRun();
		pic_run.setText("");	// 设置文本为空，插入图片
		pic_run.addPicture(new FileInputStream("D:\\POI\\test.jpg"), Document.PICTURE_TYPE_PNG, "img.png", Units.toEMU(200), Units.toEMU(200));
        
		// 创建表格段落
		XWPFParagraph table_paragraph = doc.createParagraph();
        XmlCursor cursor = table_paragraph.getCTP().newCursor();
        XWPFTable table = doc.insertNewTbl(cursor);// ---这个是关键

        // 设置表格宽度，第一行宽度就可以了
        table.setWidth(8500);
        
        List<User> users = new ArrayList<>();
        users.add(new User("张三", "男", "20", "1231", "成都市"));
        users.add(new User("李四", "男", "20", "", "眉山市"));
        users.add(new User("刘芫", "女", "22", "12138", "绵阳市"));
        
        for (int i = 0; i < users.size(); i++) {
        	User user = users.get(i);
			if (i == 0) {
				XWPFTableRow tableOneRowOne = table.getRow(0);//行
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "20%", "姓名");
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", user.getName());
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", "性别");
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", user.getGender());
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", "年龄");
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", user.getAge());
			}else {
				XWPFTableRow tableOneRowOne = table.createRow();
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "20%", "姓名");
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(1), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", user.getName());
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(2), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", "性别");
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(3), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", user.getGender());
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(4), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", "年龄");
				POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(5), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", user.getAge());
			}
	        if (StringUtils.isNotBlank(user.getIdcard())) {
	        	XWPFTableRow idcardRow = table.createRow();//行
	        	POIWordTextUtils.setWordCellSelfStyle(idcardRow.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "10%", "证件号码");
	        	POIWordTextUtils.setWordCellSelfStyle(idcardRow.getCell(1), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "45%", user.getIdcard());
	            idcardRow.getCell(1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
	            for (int j = 2; j < 6; j++) {
	            	idcardRow.getCell(j).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
				}
			}
	        XWPFTableRow addressRow = table.createRow();//行
	        POIWordTextUtils.setWordCellSelfStyle(addressRow.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "10%", "居住地址");
	        POIWordTextUtils.setWordCellSelfStyle(addressRow.getCell(1), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "45%", user.getAddress());
            addressRow.getCell(1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            for (int j = 2; j < 6; j++) {
            	addressRow.getCell(j).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}

        // 表格合并单元格测试
        XWPFParagraph paragraph4 = doc.createParagraph();
        XmlCursor cursor1 = paragraph4.getCTP().newCursor();
        XWPFTable tableOne1 = doc.insertNewTbl(cursor1);// ---这个是关键
        tableOne1.setWidth(8500);
        
        // 表格第一行，对于每个列，必须使用createCell()，而不是getCell()，因为第一行嘛，肯定是属于创建的，没有create哪里来的get呢
        XWPFTableRow tableOneRowOne = tableOne1.getRow(0);//行
        POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "20%", "序号");
        POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", "公司名称(英文)");
        POIWordTextUtils.setWordCellSelfStyle(tableOneRowOne.createCell(), "微软雅黑", "9", 0, "left", "top", "#000000", "#FFFFFF", "40%", "公司名称(中文)");

        XWPFTableRow tableOneRowTwo1 = tableOne1.createRow();//行
        POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo1.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "10%", "测试");
        POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo1.getCell(1), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "45%", "测试");
        tableOneRowTwo1.getCell(1).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
        tableOneRowTwo1.getCell(2).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
        
        for (int i = 0; i < 10; i ++) {
            // 表格第二行
            XWPFTableRow tableOneRowTwo = tableOne1.createRow();//行
            POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo.getCell(0), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "10%", "一行一列");
            POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo.getCell(1), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "45%", "一行一列");
            POIWordTextUtils.setWordCellSelfStyle(tableOneRowTwo.getCell(2), "微软雅黑", "9", 0, "left", "top", "#000000", "#ffffff", "45%", "一行一列");
        }
        
        
        
        // 横着合并单元格  -----------------------------
//        XWPFTableCell cell72 = tableOne1.getRow(6).getCell(1);  // 第7行的第2列
//        XWPFTableCell cell73 = tableOne1.getRow(6).getCell(2);  // 第7行的第3列
//        
//        cell72.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
//        cell73.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
        
//        WorderToNewWordUtils.mergeCellHorizontally(tableOne1, 7, 1, 2);
//        WorderToNewWordUtils.mergeCellHorizontally(tableOne1, 8, 1, 2);
//        WorderToNewWordUtils.mergeCellHorizontally(tableOne1, 9, 1, 2);
        // 竖着合并单元格  -----------------------------
        // XWPFTableCell cell1 = tableOne.getRow(1).getCell(0);  // 第2行第1列  第1行是表头
//        XWPFTableCell cell2 = tableOne1.getRow(2).getCell(0);  // 第3行第1列
//        XWPFTableCell cell3 = tableOne1.getRow(3).getCell(0);  // 第4行第1列
//        XWPFTableCell cell4 = tableOne1.getRow(4).getCell(0);  // 第5行第1列
//
//        // cell1.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
//        cell2.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
//        cell3.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
//        cell4.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
        
        // WorderToNewWordUtils.mergeCellVertically(tableOne1, 1, 7, 9);
        POIWordTextUtils.mergeCell(tableOne1, 1, 7, 2, 9);
	}
	
}

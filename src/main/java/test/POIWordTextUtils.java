package test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
 
/**
 * @Description:POI处理Word文本相关内容工具类（包括文本和表格）
 * @author chenran
 * @date 2020年7月1日
 */
public class POIWordTextUtils {
	
	/**
	 * 设置表格单元格格式
	 * @param cell		// 单元格
	 * @param fontName	// 字体名称
	 * @param fontSize	// 字体大小
	 * @param fontBlod	// 字体加粗
	 * @param alignment	// 水平位置
	 * @param vertical	// 垂直高度
	 * @param fontColor	// 字体颜色
	 * @param bgColor	// 背景颜色
	 * @param cellWidth	// 单元格宽度
	 * @param content	// 内容
	 */
	public static void setWordCellSelfStyle(XWPFTableCell cell, String fontName, String fontSize, int fontBlod,
			String alignment, String vertical, String fontColor, String bgColor, String cellWidth, String content) {

		// poi对字体大小设置特殊，不支持小数，但对原word字体大小做了乘2处理
		BigInteger bFontSize = new BigInteger("24");
		if (fontSize != null && !fontSize.equals("")) {
			// poi对字体大小设置特殊，不支持小数，但对原word字体大小做了乘2处理
			BigDecimal fontSizeBD = new BigDecimal(fontSize);
			fontSizeBD = new BigDecimal("2").multiply(fontSizeBD);
			fontSizeBD = fontSizeBD.setScale(0, BigDecimal.ROUND_HALF_UP);// 这里取整
			bFontSize = new BigInteger(fontSizeBD.toString());// 字体大小
		}

		// 设置单元格宽度
		cell.setWidth(cellWidth);

		// =====获取单元格
		CTTc tc = cell.getCTTc();
		// ====tcPr开始====》》》》
		CTTcPr tcPr = tc.getTcPr();// 获取单元格里的<w:tcPr>
		if (tcPr == null) {// 没有<w:tcPr>，创建
			tcPr = tc.addNewTcPr();
		}

		// --vjc开始-->>
		CTVerticalJc vjc = tcPr.getVAlign();// 获取<w:tcPr> 的<w:vAlign w:val="center"/>
		if (vjc == null) {// 没有<w:w:vAlign/>，创建
			vjc = tcPr.addNewVAlign();
		}
		// 设置单元格对齐方式
		vjc.setVal(vertical.equals("top") ? STVerticalJc.TOP
				: vertical.equals("bottom") ? STVerticalJc.BOTTOM : STVerticalJc.CENTER); // 垂直对齐

		CTShd shd = tcPr.getShd();// 获取<w:tcPr>里的<w:shd w:val="clear" w:color="auto" w:fill="C00000"/>
		if (shd == null) {// 没有<w:shd>，创建
			shd = tcPr.addNewShd();
		}
		// 设置背景颜色
		shd.setFill(bgColor.substring(1));
		// 《《《《====tcPr结束====

		// ====p开始====》》》》
		CTP p = tc.getPList().get(0);// 获取单元格里的<w:p w:rsidR="00C36068" w:rsidRPr="00B705A0" w:rsidRDefault="00C36068"
										// w:rsidP="00C36068">

		// ---ppr开始--->>>
		CTPPr ppr = p.getPPr();// 获取<w:p>里的<w:pPr>
		if (ppr == null) {// 没有<w:pPr>，创建
			ppr = p.addNewPPr();
		}
		// --jc开始-->>
		CTJc jc = ppr.getJc();// 获取<w:pPr>里的<w:jc w:val="left"/>
		if (jc == null) {// 没有<w:jc/>，创建
			jc = ppr.addNewJc();
		}
		// 设置单元格对齐方式
		jc.setVal(alignment.equals("left") ? STJc.LEFT : alignment.equals("right") ? STJc.RIGHT : STJc.CENTER); // 水平对齐
		// <<--jc结束--
		// --pRpr开始-->>
		CTParaRPr pRpr = ppr.getRPr(); // 获取<w:pPr>里的<w:rPr>
		if (pRpr == null) {// 没有<w:rPr>，创建
			pRpr = ppr.addNewRPr();
		}
		CTFonts pfont = pRpr.getRFonts();// 获取<w:rPr>里的<w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/>
		if (pfont == null) {// 没有<w:rPr>，创建
			pfont = pRpr.addNewRFonts();
		}
		// 设置字体
		pfont.setAscii(fontName);
		pfont.setEastAsia(fontName);
		pfont.setHAnsi(fontName);

		CTOnOff pb = pRpr.getB();// 获取<w:rPr>里的<w:b/>
		if (pb == null) {// 没有<w:b/>，创建
			pb = pRpr.addNewB();
		}
		// 设置字体是否加粗
		pb.setVal(fontBlod == 1 ? STOnOff.ON : STOnOff.OFF);

		CTHpsMeasure psz = pRpr.getSz();// 获取<w:rPr>里的<w:sz w:val="32"/>
		if (psz == null) {// 没有<w:sz w:val="32"/>，创建
			psz = pRpr.addNewSz();
		}
		// 设置单元格字体大小
		psz.setVal(bFontSize);
		CTHpsMeasure pszCs = pRpr.getSzCs();// 获取<w:rPr>里的<w:szCs w:val="32"/>
		if (pszCs == null) {// 没有<w:szCs w:val="32"/>，创建
			pszCs = pRpr.addNewSzCs();
		}
		// 设置单元格字体大小
		pszCs.setVal(bFontSize);
		// <<--pRpr结束--
		// <<<---ppr结束---

		// ---r开始--->>>
		List<CTR> rlist = p.getRList(); // 获取<w:p>里的<w:r w:rsidRPr="00B705A0">
		CTR r = null;
		if (rlist != null && rlist.size() > 0) {// 获取第一个<w:r>
			r = rlist.get(0);
		} else {// 没有<w:r>，创建
			r = p.addNewR();
		}
		// --rpr开始-->>
		CTRPr rpr = r.getRPr();// 获取<w:r w:rsidRPr="00B705A0">里的<w:rPr>
		if (rpr == null) {// 没有<w:rPr>，创建
			rpr = r.addNewRPr();
		}
		// ->-
		CTFonts font = rpr.getRFonts();// 获取<w:rPr>里的<w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"
										// w:hint="eastAsia"/>
		if (font == null) {// 没有<w:rFonts>，创建
			font = rpr.addNewRFonts();
		}
		// 设置字体
		font.setAscii(fontName);
		font.setEastAsia(fontName);
		font.setHAnsi(fontName);

		CTOnOff b = rpr.getB();// 获取<w:rPr>里的<w:b/>
		if (b == null) {// 没有<w:b/>，创建
			b = rpr.addNewB();
		}
		// 设置字体是否加粗
		b.setVal(fontBlod == 1 ? STOnOff.ON : STOnOff.OFF);
		CTColor color = rpr.getColor();// 获取<w:rPr>里的<w:color w:val="FFFFFF" w:themeColor="background1"/>
		if (color == null) {// 没有<w:color>，创建
			color = rpr.addNewColor();
		}
		// 设置字体颜色
		if (content.contains("↓")) {
			color.setVal("43CD80");
		} else if (content.contains("↑")) {
			color.setVal("943634");
		} else {
			color.setVal(fontColor.substring(1));
		}
		CTHpsMeasure sz = rpr.getSz();
		if (sz == null) {
			sz = rpr.addNewSz();
		}
		sz.setVal(bFontSize);
		CTHpsMeasure szCs = rpr.getSzCs();
		if (szCs == null) {
			szCs = rpr.addNewSz();
		}
		szCs.setVal(bFontSize);
		// -<-
		// <<--rpr结束--
		List<CTText> tlist = r.getTList();
		CTText t = null;
		if (tlist != null && tlist.size() > 0) {// 获取第一个<w:r>
			t = tlist.get(0);
		} else {// 没有<w:r>，创建
			t = r.addNewT();
		}
		t.setStringValue(content);
		// <<<---r结束---
	}
	
	/**
     * 垂直合并，也就是合并行（针对一列）
     * @param table		表格
     * @param col		合并列索引
     * @param fromRow	起始行索引
     * @param toRow		结束行索引
     */
    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for(int rowIndex = fromRow; rowIndex <= toRow; rowIndex++){
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if(rowIndex == fromRow){
                vmerge.setVal(STMerge.RESTART);
            } else {
                vmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setVMerge(vmerge);
            } else {
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setVMerge(vmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }
    
    /**
     * 水平合并，也就是合并列（针对一行）
     * @param table		表格
     * @param row		合并行索引
     * @param fromCol	起始列索引
     * @param toCol		结束列索引
     */
    public static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) { 
    	
    	for(int colIndex = fromCol; colIndex <= toCol; colIndex++){
    		XWPFTableCell cell = table.getRow(row).getCell(colIndex);
    		if (colIndex == fromCol) {
    			cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			}else {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
        }
 
    }
    
    /**
     * 单元格合并，表示将单元格从第（fromCol+1）列、第（fromRow+1）行到第（toCol+1）列、（toRow+1）行进行合并 
     * @param table		表格
     * @param fromCol	起始列索引
     * @param fromRow	起始行索引
     * @param toCol		结束列索引
     * @param toRow		结束行索引
     */
    public static void mergeCell(XWPFTable table, int fromCol, int fromRow, int toCol, int toRow) { 
    	
    	for(int rowIndex = fromRow; rowIndex <= toRow; rowIndex++){
    		POIWordTextUtils.mergeCellHorizontally(table, rowIndex, fromCol, toCol);
    	}
    	POIWordTextUtils.mergeCellVertically(table, fromCol, fromRow, toRow);
 
    }
 
    /**
     * 替换段落中文本和图片信息
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @throws IOException 
     * @throws FileNotFoundException 
     * @throws InvalidFormatException 
     */
    public static void changeText(XWPFDocument document, Map<String, Object> textMap) throws InvalidFormatException, FileNotFoundException, IOException{
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
 
        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if(checkText(text)){
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    if (ob != null) {
                    	// 说明是文本
                    	if (ob instanceof String){
                            run.setText((String)ob,0);
                        // 说明是图片
                        }else if (ob instanceof Map){
                            run.setText("",0);
                            Map pic = (Map)ob;
                            int width = Integer.parseInt(pic.get("width").toString());
                            int height = Integer.parseInt(pic.get("height").toString());
                            int picType = getPictureType(pic.get("type").toString());
                            String imgPath =  pic.get("path").toString();
                            try {
                            	run.addPicture(new FileInputStream(imgPath), picType, "img.png", Units.toEMU(width), Units.toEMU(height));
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
					}
                }
            }
        }
    }
 
    /**
     * 替换表格对象方法
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @param mapList 需要动态生成的内容
     */
    public static void changeTable(XWPFDocument document, Map<String, Object> textMap){
        
    	//获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        
        //循环所有需要进行替换的文本，进行替换
        for (int i = 0; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);
            if(checkText(table.getText())){
                List<XWPFTableRow> rows = table.getRows();
                //遍历表格，并替换模板
                eachTable(document,rows, textMap);
            }
        }
    }
    
    /**
     * 遍历表格替换内容
     * @param rows 表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(XWPFDocument document,List<XWPFTableRow> rows ,Map<String, Object> textMap){
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            Object ob = changeValue(run.toString(), textMap);
                            if (ob != null) {
                            	if (ob instanceof String){
                                    run.setText((String)ob,0);
                                }else if (ob instanceof Map){
                                    run.setText("",0);
                                    Map pic = (Map)ob;
                                    int width = Integer.parseInt(pic.get("width").toString());
                                    int height = Integer.parseInt(pic.get("height").toString());
                                    int picType = getPictureType(pic.get("type").toString());
                                    String imgPath =  pic.get("path").toString();
                                    try {
                                    	run.addPicture(new FileInputStream(imgPath), picType, "img.png", Units.toEMU(width), Units.toEMU(height));
                                    } catch (Exception e) {
                                        e.printStackTrace();
                                    }
                                }
							}
                        }
                    }
                }
            }
        }
    }
 
    /**
     * 判断文本中时候包含$
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        if(text.indexOf("$")!= -1){
            check = true;
        }
        return check;
    }
 
    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static Object changeValue(String value, Map<String, Object> textMap){
        Set<Entry<String, Object>> textSets = textMap.entrySet();
        Object valu = null;
        for (Entry<String, Object> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if(value.indexOf(key)!= -1){
                valu = textSet.getValue();
            }
        }
        return valu;
    }
 
    /**
     * 根据图片类型，取得对应的图片类型代码
     * @param picType
     * @return int
     */
    private static int getPictureType(String picType){
        int res = XWPFDocument.PICTURE_TYPE_PICT;
        if(picType != null){
            if(picType.equalsIgnoreCase("png")){
                res = XWPFDocument.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){
                res = XWPFDocument.PICTURE_TYPE_DIB;
            }else if(picType.equalsIgnoreCase("emf")){
                res = XWPFDocument.PICTURE_TYPE_EMF;
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){
                res = XWPFDocument.PICTURE_TYPE_JPEG;
            }else if(picType.equalsIgnoreCase("wmf")){
                res = XWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }
     
}

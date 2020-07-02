package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * @Description:表格内容替换Demo（包括文本、图片和表格相关操作）
 * @author chenran
 * @date 2020年7月1日
 */
public class TableReplaceDemo {

	public static void main(String[] args) throws Exception {
		
		System.out.println("POI表格内容替换测试");

		final String returnurl = "D:\\POI\\table_result.docx"; // 结果文件

		final String templateurl = "D:\\POI\\table.docx"; // 模板文件

		InputStream is = new FileInputStream(new File(templateurl));
		XWPFDocument doc = new XWPFDocument(is);
		replaceTable(doc);

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

	public static void replaceTable(XWPFDocument doc) throws Exception {
		// 需要进行文本替换的信息
		Map<String, Object> data = new HashMap<String, Object>();
		data.put("${date}", "2018-03-06");
		data.put("${name}", "东方明珠");
		data.put("${address}", "上海黄浦江附近");
		data.put("${communityName}", "社区名字");
		data.put("${safetyCode}", "东方社区");
		data.put("${picture2}", "图片2");
		data.put("${picture3}", "图片3");
		data.put("${buildingValue2}", "漫展提示");
		data.put("${patrolPhoto1}", "其他图片1");
		data.put("${patrolPhoto2}", "其他图片2");
		data.put("${buildingValue3}", "中国标语");

		// 图片，如果是多个图片，就新建多个map
		Map<String, Object> picture1 = new HashMap<String, Object>();
		picture1.put("width", 100);
		picture1.put("height", 150);
		picture1.put("type", "jpg");
		picture1.put("path", "D:\\POI\\test.jpg");
		data.put("${picture1}", picture1);

		// 第一个动态生成的数据列表
		List<String[]> list01 = new ArrayList<String[]>();
		list01.add(new String[] { "A", "美女好看" });
		list01.add(new String[] { "A", "美女好多" });
		list01.add(new String[] { "B", "漫展人太多" });
		list01.add(new String[] { "C", "妹子穿的很清凉" });

		// 第二个动态生成的数据列表
		List<String> list02 = new ArrayList<String>();
		list02.add("1、民主");
		list02.add("2、富强");
		list02.add("3、文明");
		list02.add("4、和谐");
		
		POIWordTextUtils.changeText(doc, data);
		
		POIWordTextUtils.changeTable(doc, data);
		
		// 动态生成表格内容，属于定制化需求
		List<XWPFTable> tables = doc.getTables();
        //操作word中的表格
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            XWPFTable table = tables.get(i);
            //第二个表格使用daList，插入数据
            if (null != list01 && 0 < list01.size() && i == 1){
                insertTable(table, null,list01,2);
                // 创建列并且合并
                table.getRow(0).createCell();
                POIWordTextUtils.mergeCellHorizontally(table, 0, 0, 1);
                // 合并行
                List<Integer[]> indexList = startEnd(list01);
                for (int c=0;c<indexList.size();c++){
                    //合并行
                	POIWordTextUtils.mergeCellVertically(table,0,indexList.get(c)[0]+1,indexList.get(c)[1]+1);
                }
            }
            //第四个表格使用tableList，插入数据
            if (null != list02 && 0 < list02.size() && i == 3){
                insertTable(table, list02,null,4);
            }
        }
		
	}
	
	/**
     * 为表格插入数据，行数不够添加新行
     * @param table 需要插入数据的表格
     * @param tableList 第四个表格的插入数据
     * @param daList 第二个表格的插入数据
     * @param type 表格类型：1-第一个表格 2-第二个表格 3-第三个表格 4-第四个表格
     */
    public static void insertTable(XWPFTable table, List<String> tableList,List<String[]> daList,Integer type){
        if (2 == type){
            //创建行和创建需要的列
            for(int i = 0; i < daList.size(); i++){
                XWPFTableRow row = table.createRow();
                row.createCell();//创建第二列
            }
            //创建行，根据需要插入的数据添加新行，不处理表头
            for(int i = 0; i < daList.size(); i++){
                List<XWPFTableCell> cells = table.getRow(i+1).getTableCells();
                for(int j = 0; j < cells.size(); j++){
                    XWPFTableCell cell02 = cells.get(j);
                    cell02.setText(daList.get(i)[j]);
                }
            }
        }else if (4 == type){
            //插入表头下面第一行的数据
            for(int i = 0; i < tableList.size(); i++){
                XWPFTableRow row = table.createRow();
                List<XWPFTableCell> cells = row.getTableCells();
                cells.get(0).setText(tableList.get(i));
            }
        }
    }
    
    /**
     * 获取需要合并单元格的下标
     * @return
     */
    public static List<Integer[]> startEnd(List<String[]> daList){
        List<Integer[]> indexList = new ArrayList<Integer[]>();
 
        List<String> list = new ArrayList<String>();
        for (int i=0;i<daList.size();i++){
            list.add(daList.get(i)[0]);
        }
        Map<Object, Integer> tm = new HashMap<Object, Integer>();
        for (int i=0;i<daList.size();i++){
            if (!tm.containsKey(daList.get(i)[0])) {
                tm.put(daList.get(i)[0], 1);
            } else {
                int count = tm.get(daList.get(i)[0]) + 1;
                tm.put(daList.get(i)[0], count);
            }
        }
        for (Map.Entry<Object, Integer> entry : tm.entrySet()) {
            String key = entry.getKey().toString();
            // String value = entry.getValue().toString();
            if (list.indexOf(key) != (-1)){
                Integer[] index = new Integer[2];
                index[0] = list.indexOf(key);
                index[1] = list.lastIndexOf(key);
                indexList.add(index);
            }
        }
        return indexList;
    }
}

package test;

import java.util.Date;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.NumberToTextConverter;

import com.alibaba.fastjson.util.TypeUtils;

/**
 * @Description:POI处理Excel相关工具类
 * @author chenran
 * @date 2020年7月1日
 */
public class POIExcelUtils {

	/**
     * 从单元格中提取字符串(trimToNull)
     *
     * @param cell 单元格
     * @return
     */
    public static final String getCellStringValue(Cell cell) {
        return getCellStringValue(cell, true);
    }

    /**
     * 从单元格中提取字符串
     *
     * @param cell       单元格
     * @param trimToNull 是否 trimToNull
     * @return
     */
    public static final String getCellStringValue(Cell cell, boolean trimToNull) {

        if (cell == null) {
            return "";
        }
        CellType type = cell.getCellType();
        String cellValue;
        switch (type) {

            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                // short format = cell.getCellStyle().getDataFormat();

                /**
                 * @BuiltinFormats 具体格式参见此类
                 */
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = DateFormatUtils.format(DateUtil.getJavaDate(cell.getNumericCellValue()), "yyyy-MM-dd HH:mm:ss");
                } else {
                    cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                }

                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case FORMULA:
                cellValue = cell.getCellFormula();
                break;
            default:
                cellValue = "";
                break;
        }
        if (trimToNull) {
            return StringUtils.trimToNull(cellValue);
        } else {
            return cellValue;
        }
    }

    /**
     * 从单元格中提取时间
     *
     * @param cell
     * @return
     */
    public static final Date getCellDateValue(Cell cell) {

        if (cell == null) {
            return null;
        }

        CellType type = cell.getCellType();

        if (type.equals(CellType.NUMERIC)) {
            try {
                return DateUtil.getJavaDate(cell.getNumericCellValue());
            } catch (RuntimeException e) {

            }
        }

        try {
            return TypeUtils.castToDate(getCellStringValue(cell));
        } catch (RuntimeException e) {
            return null;
        }

    }
}

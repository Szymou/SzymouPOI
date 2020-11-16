package SzymouPOI.utils;

import com.alibaba.fastjson.JSON;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * create by cyh
 * POI导入Excel文件并读取工具
 * 2020年11月13日 13:51:24
 */
public class SzymouPoiUtil {
    Logger logger = LoggerFactory.getLogger(this.getClass());

    public <T> List<T> readExcel(File file, Class<T> tClass, String... attr) throws IOException {

        // 判断是否为空
        if (!file.exists()) {
            return null;
        }
        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);//读取第一个sheet
        inputStream.close();//读取到文件内容，可以关掉文件流了。
        int rows = sheet.getLastRowNum();//共有几行数据
//        int cols = colNum;//返回实体类的字段数，即Excel表列数
        int cols = attr.length;//返回实体类的字段数，即Excel表列数

        XSSFRow row = null;
        XSSFCell cell = null;
        String content = "";
        List<Map<String, Object>> mapList = new ArrayList<>();
        List<T> tList = new ArrayList<>();
        for (int i = 1; i <= rows; i++) {
            logger.info("读取第{}行。", i);//从0开始
            row = sheet.getRow(i);
            if (null != row) {
                Map<String, Object> map = new HashMap<>();
                for (int c = 0; c < cols; c++) {
                    logger.info("读取第{}行第{}列。", i, c);
                    cell = row.getCell(c);
                    content = getCellValue(cell);
                    map.put(attr[c], content);
                }
                mapList.add(map);
                T t = JSON.parseObject(JSON.toJSONString(map), tClass);
                tList.add(t);
            }
        }

        return tList;
    }


    private String getCellValue(XSSFCell cell){
        if(null != cell){
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getRichStringCellValue().getString();
                case NUMERIC:
                    return (new Double(cell.getNumericCellValue())).intValue() + "";
                default:
                    return "";
            }
        }

        return "";
    }
}

package com.hm.util.excel;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 读取excel通用类
 *
 * @author xwf
 * @version 2017-4-12
 */
public class ExcelReaderUtil {
    private static Logger log = LoggerFactory.getLogger(ExcelReaderUtil.class);

    /**
     * 读取excel内容
     *
     * @param inputStream excel输入流
     * @param fileName    文件名
     * @return
     */
    public List<Map<Integer, Object>> readExcelContent(InputStream inputStream, String fileName) {
        if (StringUtils.isBlank(fileName)) {
            return null;
        }
        String ext = FilenameUtils.getExtension(fileName);
        boolean isXlsx = ext.equalsIgnoreCase("xlsx");
        return readExcelContent(inputStream, isXlsx, 0);
    }

    /**
     * 读取excel内容
     *
     * @param inputStream excel输入流
     * @param isXlsx      是否为xlsx格式2007版本
     * @param sheetIndex  工作表下标
     * @return
     */
    public List<Map<Integer, Object>> readExcelContent(InputStream inputStream, boolean isXlsx, int sheetIndex) {
        if (inputStream == null) {
            return null;
        }
        Workbook workbook = null;
        try {
            if (isXlsx) {
                workbook = new XSSFWorkbook(inputStream);
            } else {
                workbook = new HSSFWorkbook(inputStream);
            }
        } catch (IOException e) {
            log.warn("read excel file failed!", e);
            return null;
        } finally {
            IOUtils.closeQuietly(inputStream);
        }

        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) {
            return null;
        }

        List<Map<Integer, Object>> list = new ArrayList<Map<Integer, Object>>();
        int colNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //需要组装的列数据
        List<Integer> keys = new ArrayList<>();
        //进入表格标记
        Boolean headFlag = false;
        //进入表内容标记
        Boolean bodyFlag = false;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            Map<Integer, Object> c = new HashMap<Integer, Object>();
            Boolean flag = false;
            for (int j = 0; j < colNum; j++) {
                CellStyle cellStyle = row.getCell(j).getCellStyle();
                //获取背景色
                short fillForegroundColor = cellStyle.getFillForegroundColor();
                short fillBackgroundColor = cellStyle.getFillBackgroundColor();
                if (row.getCell(j) != null) {
                    row.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
                    //获取单元格内容
                    String cellValue = row.getCell(j).getStringCellValue();
                    //进入标记并且没有背景色则为需要的列
                    if(headFlag && fillForegroundColor==64){
                        keys.add(j);
                    }
                    if (StringUtils.isNotEmpty(cellValue)) {
                        flag = true;
                    }
                    //表格标记
                    if("包号".equals(cellValue)){
                        //包号列必要
                        keys.add(j);
                        headFlag = true;
                    }
                    if(headFlag){
                        System.out.println(fillForegroundColor+":"+cellValue);
                        System.out.println(fillBackgroundColor+":"+cellValue);
                    }
                }
                if(headFlag || bodyFlag){
                    if(keys.contains(j)){
                        c.put(j, getCellValue(row.getCell(j)));
                    }
                }else{
                    c.put(j, getCellValue(row.getCell(j)));
                }
            }
            //打开另外一个标记
            if(headFlag){
                bodyFlag = true;
            }
            //关闭表投标记
            headFlag = false;
            if (flag) {
                list.add(c);
            }
        }
        return list;
    }

    /**
     * 根据excel类型取值
     *
     * @param cell
     * @return
     */
    private Object getCellValue(Cell cell) {// 获取单元格数据内容为字符串类型的数据
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue().trim();
            case Cell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat TIME_FORMATTER = new SimpleDateFormat("yyyy-MM-dd");
                    return TIME_FORMATTER.format(cell.getDateCellValue());
                } else {

                }
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            default:
                return "";
        }
    }


    public String toString(Object obj) {
        if (obj == null) {
            return null;
        }
        return obj.toString();
    }
}
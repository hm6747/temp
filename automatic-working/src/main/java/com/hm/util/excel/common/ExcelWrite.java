package com.hm.util.excel.common;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2018/5/29 0029.
 */
public class ExcelWrite {

    /**
     * 导出EXCEL
     *
     * @param list     需要导出的集合:集合中可以为map，也可以为Dto
     * @return
     */
    public static String exp(String path,List<Object> list,String [] fields,HttpServletResponse response,String fileName) {
        int startRow = 2; // 开始写入excel的开始行号
        String date = new SimpleDateFormat("yyyyMMdd").format(new Date());
        OutputStream os = null;
        InputStream is = null;
        try {
            is = new FileInputStream(path);
            Workbook work = new HSSFWorkbook(is);
            // 得到excel的第0张表
            Sheet sheet = work.getSheetAt(0);
            // 得到第1行的第一个单元格的样式
            CellStyle cellStyle = createBorderedStyle(work);
            Row row = null;
            Cell cell = null;
            Object rDto = null;
            Map<String,String> map = null;
            // 得到行，并填充数据和表格样式
            for (int i = 0; i < list.size(); i++) {
                rDto = list.get(i);
                map = (Map) rDto;
                row = sheet.createRow(startRow++);
                row.setHeight((short) 350);
                for (int j = 0; j < fields.length + 1; j++) {
                    cell = row.createCell(j);
                    if (j == 0)
                        cell.setCellValue(i + 1);
                    else
                        cell.setCellValue(map.get(fields[j - 1]));
                    cell.setCellStyle(cellStyle);// 填充样式
                }
            }
            // 输出工作簿: 这里使用的是 response 的输出流，如果将该输出流换为普通的文件输出流则可以将生成的文档写入磁盘等
            os = response.getOutputStream();
            response.setContentType("application/ms-excel,charset=gbk");
            response.setHeader("Content-disposition", "attachment;filename="
                    + URLEncoder.encode(fileName + ".xls", "utf-8"));
            work.write(os); // 将工作簿进行输出
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            close(is, os);
        }
        return null;
    }

    // 单元格样式
    protected static CellStyle createBorderedStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        return style;
    }

    // 单元格样式:获取的第二行的第一个单元格的样式
    protected static CellStyle createGetCellStyle(Workbook work) {
        Sheet sheet = work.getSheetAt(0);
        Row rowCellStyle = sheet.getRow(1);
        CellStyle cellStyle = rowCellStyle.getCell(0).getCellStyle();
        return cellStyle;
    }

    // 关闭IO
    protected static void close(InputStream is, OutputStream os) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        if (os != null) {
            try {
                os.flush();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

}

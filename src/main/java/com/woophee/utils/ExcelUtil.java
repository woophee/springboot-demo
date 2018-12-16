package com.woophee.utils;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

    /**
     * 获取数据
     * @param file
     * @return
     * @throws Exception
     */
    public static List<List<String>> readExcel(File file) throws Exception {

        // 创建输入流，读取Excel
        InputStream is = new FileInputStream(file.getAbsolutePath());
        // jxl提供的Workbook类
        Workbook wb = Workbook.getWorkbook(is);
        // 只有一个sheet,直接处理
        //创建一个Sheet对象
        Sheet sheet = wb.getSheet(0);
        // 得到所有的行数
        int rows = sheet.getRows();
        // 所有的数据
        List<List<String>> allData = new ArrayList<List<String>>();
        // 越过第一行 它是列名称
        for (int j = 0; j < rows; j++) {

            List<String> oneData = new ArrayList<String>();
            // 得到每一行的单元格的数据
            Cell[] cells = sheet.getRow(j);
            for (int k = 0; k < cells.length; k++) {

                oneData.add(cells[k].getContents().trim());
            }
            // 存储每一条数据
            allData.add(oneData);
            // 打印出每一条数据
//            System.out.println(oneData);

        }
        return allData;

    }

    public static void makeExcel(List<List<String>> result) throws IOException {

        //工作表名后面的数字，如表1，表2
        int sheetNum = 0;
        //记录总行数
        int rownum = 0;
        //记录每个sheet的行数
        int tempnum = 0;
        //分页条数达到此条数则创建工作表
        int page = 60000;

        //第一步，创建一个workbook对应一个excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        //第二部，在workbook中创建一个sheet对应excel中的sheet

        while(true) {
            HSSFSheet sheet = workbook.createSheet("sheet" + (++sheetNum));
            if (result != null) {
                tempnum = 0;
                for (int i = rownum; i < result.size(); i++) {
                    List<String> oneData = result.get(i);
                    HSSFRow row1 = sheet.createRow(tempnum++);
                    rownum++;
                    for (int j = 0; j < oneData.size(); j++) {

                        //创建单元格设值
                        row1.createCell(j).setCellValue(oneData.get(j));
                    }
                    if (rownum % page == 0) {
                        break;
                    }
                }
            }
            if(rownum == result.size()){
                break;
            }
        }

        long timestamp = System.currentTimeMillis();
        //将文件保存到指定的位置
        FileOutputStream fos = new FileOutputStream("result" + timestamp + ".xls");
        workbook.write(fos);
        fos.close();
    }


}

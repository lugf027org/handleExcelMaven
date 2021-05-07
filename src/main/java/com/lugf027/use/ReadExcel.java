package com.lugf027.use;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    public static void main(String[] args) {
        ReadExcel obj = new ReadExcel();
        // 此处为我创建Excel路径：E:/zhanhj/studysrc/jxl下
        File file = new File(System.getProperty("user.dir") + "\\readExcel.xls");
        List excelDataList = obj.readExcel(file);

//        System.out.println("list中的数据打印出来");
//        for (int i = 0; i < excelDataList.size(); i++) {
//            List list = (List) excelDataList.get(i);
//            for (int j = 0; j < list.size(); j++) {
//                System.out.print(String.valueOf(j) + list.get(j));
//            }
//            System.out.println();
//        }


        Map<String, String> cityCodeMap = new HashMap<String, String>();
        for (int i = 0; i < excelDataList.size(); i++) {
            List list = (List) excelDataList.get(i);
            if (list.size() == 4) {
                String key = list.get(2).toString();
                String value = list.get(3).toString();
                cityCodeMap.put(key, value);
            }
        }

        for (Map.Entry<String, String> entry : cityCodeMap.entrySet()) {
            String mapKey = entry.getKey();
            String mapValue = entry.getValue();
            System.out.println(mapKey + ":" + mapValue);
        }

        List<List> outListRes = new ArrayList<List>();

        for (int i = 0; i < excelDataList.size(); i++) {
            List list = (List) excelDataList.get(i);
            List innerList = new ArrayList();

            String firstStr = list.get(0).toString();
            String secondStr = list.get(1).toString();

            if (cityCodeMap.containsKey(firstStr)) {
                innerList.add(cityCodeMap.get(firstStr));
            } else {
                int braceIndex = firstStr.indexOf("(");
                String subStr = braceIndex > 0 ? firstStr.substring(0,braceIndex) : firstStr;
                if (braceIndex > 0 && cityCodeMap.containsKey(subStr)) {
                    innerList.add(cityCodeMap.get(subStr));
                } else {
                    innerList.add(firstStr + "===");
                }
            }

            if (cityCodeMap.containsKey(secondStr)) {
                innerList.add(cityCodeMap.get(secondStr));
            } else {
                int braceIndex = secondStr.indexOf("(");
                String subStr = braceIndex > 0 ? secondStr.substring(0,braceIndex) : secondStr;
                if (braceIndex > 0 && cityCodeMap.containsKey(subStr)) {
                    innerList.add(cityCodeMap.get(subStr));
                } else {
                    innerList.add(secondStr + "===");
                }
            }

            outListRes.add(innerList);
        }


        for (int i = 0; i < outListRes.size(); i++) {
            System.out.print(String.valueOf(i) + "\t");
            List list = (List) outListRes.get(i);
            for (int j = 0; j < list.size(); j++) {
                System.out.print(String.valueOf(j) + "\t" + list.get(j));
            }
            System.out.println();
        }

        writeExcel(outListRes, 2, System.getProperty("user.dir") + "\\out1.xlsx");

    }

    public static void writeExcel(List<List> dataListToWrite, int cloumnCount, String finalXlsxPath) {
        OutputStream out = null;
        try {
            // 获取总列数
            int columnNumCount = cloumnCount;
            // 读取Excel文档
            File finalXlsxFile = new File(finalXlsxPath);
            org.apache.poi.ss.usermodel.Workbook workBook = getWorkbok(finalXlsxFile);
            // sheet 对应一个工作页
            org.apache.poi.ss.usermodel.Sheet sheet = workBook.getSheetAt(0);
            /**
             * 删除原有数据，除了属性列
             */
            int rowNumber = sheet.getLastRowNum();    // 第一行从0开始算
            System.out.println("原始数据总行数，除属性列：" + rowNumber);
            for (int i = 1; i <= rowNumber; i++) {
                Row row = sheet.getRow(i);
                sheet.removeRow(row);
            }
            // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(finalXlsxPath);
            workBook.write(out);
            /**
             * 往Excel中写新数据
             */
            for (int j = 0; j < dataListToWrite.size(); j++) {
                List innerList = dataListToWrite.get(j);

                // 创建一行
                Row row = sheet.createRow(j);

                for (int k = 0; k < innerList.size(); ++k) {
                    // 在一行内循环
                    Cell first = row.createCell(k);
                    first.setCellValue(innerList.get(k).toString());
                }
            }
            // 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(finalXlsxPath);
            workBook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.flush();
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("数据导出成功");
    }

    /**
     * 判断Excel的版本,获取Workbook
     *
     * @return
     * @throws IOException
     */
    public static org.apache.poi.ss.usermodel.Workbook getWorkbok(File file) throws IOException {
        org.apache.poi.ss.usermodel.Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if (file.getName().endsWith(EXCEL_XLS)) {     //Excel 2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    // 去读Excel的方法readExcel，该方法的入口参数为一个File对象
    public List readExcel(File file) {
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            Workbook wb = Workbook.getWorkbook(is);
            // Excel的页签数量
            int sheet_size = wb.getNumberOfSheets();
            for (int index = 0; index < sheet_size; index++) {
                List<List> outerList = new ArrayList<List>();
                // 每个页签创建一个Sheet对象
                Sheet sheet = wb.getSheet(index);
                // sheet.getRows()返回该页的总行数
                for (int i = 0; i < sheet.getRows(); i++) {
                    List innerList = new ArrayList();
                    // sheet.getColumns()返回该页的总列数
                    for (int j = 0; j < sheet.getColumns(); j++) {
                        String cellinfo = sheet.getCell(j, i).getContents();
                        if (cellinfo.isEmpty()) {
                            continue;
                        }
                        innerList.add(cellinfo);
//                        System.out.print(cellinfo);
                    }
                    outerList.add(i, innerList);
//                    System.out.println();
                }
                return outerList;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
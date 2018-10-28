//package com.stylefeng.guns.modular.system.controller;
//
//import org.apache.poi.hssf.usermodel.*;
//import org.apache.poi.hssf.util.HSSFCellUtil;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.Font;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.util.CellRangeAddress;
//
//import javax.swing.*;
//import java.io.File;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
///**
// * @Author: Albert Xiao
// * @Date: 2018/10/25 22:30
// * @Description:
// */
//public class ExcelUtil {
//    /**
//     * 功能：创建HSSFSheet工作簿
//     * @param     wb    HSSFWorkbook
//     * @param     sheetName    String
//     * @return    HSSFSheet
//     */
//    public static HSSFSheet createSheet(HSSFWorkbook wb, String sheetName){
//        HSSFSheet sheet=wb.createSheet(sheetName);
//        sheet.setDefaultColumnWidth(12);
//        sheet.setGridsPrinted(false);
//        sheet.setDisplayGridlines(false);
//        return sheet;
//    }
//
//
//
//
//    /**
//     * 功能：创建HSSFRow
//     * @param     sheet    HSSFSheet
//     * @param     rowNum    int
//     * @param     height    int
//     * @return    HSSFRow
//     */
//    public static HSSFRow createRow(HSSFSheet sheet, int rowNum, int height){
//        HSSFRow row=sheet.createRow(rowNum);
//        row.setHeight((short)height);
//        return row;
//    }
//
//
//
//    public static HSSFCell createCell0(HSSFRow row, int cellNum){
//        HSSFCell cell=row.createCell(cellNum);
//        return cell;
//    }
//
//
//    /**
//     * 功能：创建CELL
//     * @param     row        HSSFRow
//     * @param     cellNum    int
//     * @param     style    HSSFStyle
//     * @return    HSSFCell
//     */
//    public static HSSFCell createCell(HSSFRow row, int cellNum, CellStyle style){
//        HSSFCell cell=row.createCell(cellNum);
//        cell.setCellStyle(style);
//        return cell;
//    }
//
//
//
//    /**
//     * 功能：创建CellStyle样式
//     * @param     wb                HSSFWorkbook
//     * @param     backgroundColor    背景色
//     * @param     foregroundColor    前置色
//     * @param    font            字体
//     * @return    CellStyle
//     */
//    public static CellStyle createCellStyle(HSSFWorkbook wb, short backgroundColor, short foregroundColor, short halign, Font font){
//        CellStyle cs=wb.createCellStyle();
//        cs.setAlignment(halign);
//        cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        cs.setFillBackgroundColor(backgroundColor);
//        cs.setFillForegroundColor(foregroundColor);
//        cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        cs.setFont(font);
//        return cs;
//    }
//
//
//    /**
//     * 功能：创建带边框的CellStyle样式
//     * @param     wb                HSSFWorkbook
//     * @param     backgroundColor    背景色
//     * @param     foregroundColor    前置色
//     * @param    font            字体
//     * @return    CellStyle
//     */
//    public static CellStyle createBorderCellStyle(HSSFWorkbook wb,short backgroundColor,short foregroundColor,short halign,Font font){
//        CellStyle cs=wb.createCellStyle();
//        cs.setAlignment(halign);
//        cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        cs.setFillBackgroundColor(backgroundColor);
//        cs.setFillForegroundColor(foregroundColor);
//        cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        cs.setFont(font);
//        cs.setBorderLeft(CellStyle.BORDER_DASHED);
//        cs.setBorderRight(CellStyle.BORDER_DASHED);
//        cs.setBorderTop(CellStyle.BORDER_DASHED);
//        cs.setBorderBottom(CellStyle.BORDER_DASHED);
//        return cs;
//    }
//
//
//
//
//
//    /**
//     * 功能：多行多列导入到Excel并且设置标题栏格式
//     */
//    public static void writeArrayToExcel(HSSFSheet sheet,int rows,int cells,Object [][]value){
//
//        Row row[]=new HSSFRow[rows];
//        Cell cell[]=new HSSFCell[cells];
//
//        for(int i=0;i<row.length;i++){
//            row[i]=sheet.createRow(i);
//
//
//            for(int j=0;j<cell.length;j++){
//                cell[j]=row[i].createCell(j);
//                cell[j].setCellValue(convertString(value[i][j]));
//
//            }
//
//        }
//    }
//
//
//
//    /**
//     * 功能：多行多列导入到Excel并且设置标题栏格式
//     */
//    public static void writeArrayToExcel(HSSFWorkbook wb,HSSFSheet sheet,int rows,int cells,Object [][]value){
//
//        Row row[]=new HSSFRow[rows];
//        Cell cell[]=new HSSFCell[cells];
//
//
//        HSSFCellStyle ztStyle =  (HSSFCellStyle)wb.createCellStyle();
//
//        Font ztFont = wb.createFont();
//        ztFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
//        //ztFont.setItalic(true);                     // 设置字体为斜体字
//        // ztFont.setColor(Font.COLOR_RED);            // 将字体设置为“红色”
//        ztFont.setFontHeightInPoints((short)10);    // 将字体大小设置为18px
//        ztFont.setFontName("华文行楷");             // 将“华文行楷”字体应用到当前单元格上
//        // ztFont.setUnderline(Font.U_DOUBLE);
//        ztStyle.setFont(ztFont);
//
//        for(int i=0;i<row.length;i++){
//            row[i]=sheet.createRow(i);
//
//
//            for(int j=0;j<cell.length;j++){
//                cell[j]=row[i].createCell(j);
//                cell[j].setCellValue(convertString(value[i][j]));
//
//                if(i==0)
//                    cell[j].setCellStyle(ztStyle);
//
//            }
//
//        }
//    }
//
//
//
//    /**
//     * 功能：合并单元格
//     * @param     sheet        HSSFSheet
//     * @param     firstRow    int
//     * @param     lastRow        int
//     * @param     firstColumn    int
//     * @param     lastColumn    int
//     * @return    int            合并区域号码
//     */
//    public static int mergeCell(HSSFSheet sheet,int firstRow,int lastRow,int firstColumn,int lastColumn){
//        return sheet.addMergedRegion(new CellRangeAddress(firstRow,lastRow,firstColumn,lastColumn));
//    }
//
//
//
//    /**
//     * 功能：创建字体
//     * @param     wb            HSSFWorkbook
//     * @param     boldweight    short
//     * @param     color        short
//     * @return    Font
//     */
//    public static Font createFont(HSSFWorkbook wb,short boldweight,short color,short size){
//        Font font=wb.createFont();
//        font.setBoldweight(boldweight);
//        font.setColor(color);
//        font.setFontHeightInPoints(size);
//        return font;
//    }
//
//
//    /**
//     * 设置合并单元格的边框样式
//     * @param    sheet    HSSFSheet
//     * @param     ca        CellRangAddress
//     * @param     style    CellStyle
//     */
//    public static void setRegionStyle(HSSFSheet sheet, CellRangeAddress ca,CellStyle style) {
//        for (int i = ca.getFirstRow(); i <= ca.getLastRow(); i++) {
//            HSSFRow row = HSSFCellUtil.getRow(i, sheet);
//            for (int j = ca.getFirstColumn(); j <= ca.getLastColumn(); j++) {
//                HSSFCell cell = HSSFCellUtil.getCell(row, j);
//                cell.setCellStyle(style);
//            }
//        }
//    }
//
//    /**
//     * 功能：将HSSFWorkbook写入Excel文件
//     * @param     wb        HSSFWorkbook
//     */
//    public static void writeWorkbook(HSSFWorkbook wb,String fileName){
//        FileOutputStream fos=null;
//        File f=new File(fileName);
//        try {
//            fos=new FileOutputStream(f);
//            wb.write(fos);
//            int dialog = JOptionPane.showConfirmDialog(null,
//                    f.getName()+"导出成功！是否打开？",
//                    "温馨提示", JOptionPane.YES_NO_OPTION);
//            if (dialog == JOptionPane.YES_OPTION) {
//
//                Runtime.getRuntime().exec("cmd /c start \"\" \"" + fileName + "\"");
//            }
//
//
//        } catch (FileNotFoundException e) {
//            JOptionPane.showMessageDialog(null, "导入数据前请关闭工作表");
//
//        } catch ( Exception e) {
//            JOptionPane.showMessageDialog(null, "没有进行筛选");
//
//        } finally{
//            try {
//                if(fos!=null){
//                    fos.close();
//                }
//            } catch (IOException e) {
//            }
//        }
//    }
//
//
//
//    public static String convertString(Object value) {
//        if (value == null) {
//            return "";
//        } else {
//            return value.toString();
//        }
//    }
//
//
//
//
//}
//
package com.stylefeng.guns.modular.system.controller;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/*******************************************
 *
 * @Package com.cccuu.project.myUtils
 * @Author duan
 * @Date 2018/3/29 9:12
 * @Version V1.0
 *******************************************/
public class ExcelUtil {


    /**
     * 导出excel
     * @param title  导出表的标题
     * @param rowsName 导出表的列名
     * @param dataList  需要导出的数据
     * @param fileName  生成excel文件的文件名
     * @param response
     */
    public void exportExcel(String title,String[] rowsName,List<Object[]> dataList,String fileName,HttpServletResponse response) throws Exception{
        OutputStream output = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition",
                "attachment; filename="+fileName);
        response.setContentType("application/msexcel");
        this.export(title,rowsName,dataList,fileName,output);
        this.close(output);

    }



    /*
     * 导出数据
     */
    private void export(String title,String[] rowName,List<Object[]> dataList,String fileName,OutputStream out) throws Exception {
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(); // 创建工作簿对象
            HSSFSheet sheet = workbook.createSheet(title); // 创建工作表
            HSSFRow rowm = sheet.createRow(0);  // 产生表格标题行
            HSSFCell cellTiltle = rowm.createCell(0);   //创建表格标题列
            // sheet样式定义;    getColumnTopStyle();    getStyle()均为自定义方法 --在下面,可扩展
            HSSFCellStyle columnTopStyle = this.getColumnTopStyle(workbook);// 获取列头样式对象
            HSSFCellStyle style = this.getStyle(workbook); // 获取单元格样式对象
            //合并表格标题行，合并列数为列名的长度,第一个0为起始行号，第二个1为终止行号，第三个0为起始列好，第四个参数为终止列号
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, (rowName.length - 1)));
            cellTiltle.setCellStyle(columnTopStyle);    //设置标题行样式
            cellTiltle.setCellValue(title);     //设置标题行值
            int columnNum = rowName.length;     // 定义所需列数
            HSSFRow rowRowName = sheet.createRow(2); // 在索引2的位置创建行(最顶端的行开始的第二行)
            // 将列头设置到sheet的单元格中
            for (int n = 0; n < columnNum; n++) {
                HSSFCell cellRowName = rowRowName.createCell(n); // 创建列头对应个数的单元格
                cellRowName.setCellType(HSSFCell.CELL_TYPE_STRING); // 设置列头单元格的数据类型
                HSSFRichTextString text = new HSSFRichTextString(rowName[n]);
                cellRowName.setCellValue(text); // 设置列头单元格的值
                cellRowName.setCellStyle(columnTopStyle); // 设置列头单元格样式
            }

            // 将查询出的数据设置到sheet对应的单元格中
            for (int i = 0; i < dataList.size(); i++) {
                Object[] obj = dataList.get(i);   // 遍历每个对象
                HSSFRow row = sheet.createRow(i + 3);   // 创建所需的行数
                for (int j = 0; j < obj.length; j++) {
                    HSSFCell cell = null;   // 设置单元格的数据类型
                    if (j == 0) {
                        cell = row.createCell(j, HSSFCell.CELL_TYPE_NUMERIC);
                        cell.setCellValue(i + 1);
                    } else {
                        cell = row.createCell(j, HSSFCell.CELL_TYPE_STRING);
                        if (!"".equals(obj[j]) && obj[j] != null) {
                            cell.setCellValue(obj[j].toString()); // 设置单元格的值
                        }
                    }
                    cell.setCellStyle(style); // 设置单元格样式
                }
            }

            // 让列宽随着导出的列长自动适应
            for (int colNum = 0; colNum < columnNum; colNum++) {
                int columnWidth = sheet.getColumnWidth(colNum) / 256;
                for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                    HSSFRow currentRow;
                    // 当前行未被使用过
                    if (sheet.getRow(rowNum) == null) {
                        currentRow = sheet.createRow(rowNum);
                    } else {
                        currentRow = sheet.getRow(rowNum);
                    }
                    if (currentRow.getCell(colNum) != null) {
                        HSSFCell currentCell = currentRow.getCell(colNum);
                        if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                            int length = currentCell.getStringCellValue()
                                    .getBytes().length;
                            if (columnWidth < length) {
                                columnWidth = length;
                            }
                        }
                    }
                }
                if (colNum == 0) {
                    sheet.setColumnWidth(colNum, (columnWidth - 2) * 256);
                } else {
                    sheet.setColumnWidth(colNum, (columnWidth + 4) * 256);
                }
            }
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /*
     * 列头单元格样式
     */
    private HSSFCellStyle getColumnTopStyle(HSSFWorkbook workbook) {

        // 设置字体
        HSSFFont font = workbook.createFont();
        // 设置字体大小
        font.setFontHeightInPoints((short) 11);
        // 字体加粗
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 设置字体名字
        font.setFontName("Courier New");
        // 设置样式;
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置底边框;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        // 设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        // 设置左边框;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        // 设置右边框;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        // 设置右边框颜色;
        style.setRightBorderColor(HSSFColor.BLACK.index);
        // 设置顶边框;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        // 设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.BLACK.index);
        // 在样式用应用设置的字体;
        style.setFont(font);
        // 设置自动换行;
        style.setWrapText(false);
        // 设置水平对齐的样式为居中对齐;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }

    /*
     * 列数据信息单元格样式
     */
    private HSSFCellStyle getStyle(HSSFWorkbook workbook) {
        // 设置字体
        HSSFFont font = workbook.createFont();
        // 设置字体大小
        // font.setFontHeightInPoints((short)10);
        // 字体加粗
        // font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 设置字体名字
        font.setFontName("Courier New");
        // 设置样式;
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置底边框;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        // 设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        // 设置左边框;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        // 设置右边框;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        // 设置右边框颜色;
        style.setRightBorderColor(HSSFColor.BLACK.index);
        // 设置顶边框;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        // 设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.BLACK.index);
        // 在样式用应用设置的字体;
        style.setFont(font);
        // 设置自动换行;
        style.setWrapText(false);
        // 设置水平对齐的样式为居中对齐;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        return style;
    }

    /**
     * 关闭输出流
     * @param os
     */
    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
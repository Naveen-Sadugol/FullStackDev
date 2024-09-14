package com.excel.assignments;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class ExcelAssignment1 {
    public static void main(String[] args) {
        writeContents();
    }
    public static void writeContents(){
        String[] flowers={"rose","lily","lotus","sunflower","marigold","hibiscus","tulip",
        "jasmine","Diasy","Lavendar","Dahlia","Bluebell","Waterlily","Orchid","Iris",
        "Calendula","Poppy","Daffodil","Snowdrop","Geranium"};
        FileOutputStream fout = null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try{
            wb=new XSSFWorkbook();
            sh=wb.createSheet("Flowers");
            for (int i = 0; i <20 ; i++) {
                row=sh.createRow(i);
                cell=row.createCell(0);
                cell.setCellValue(flowers[i]);
            }
            fout = new FileOutputStream("E:\\Vijay Kumar A\\ExcelSheets\\Excel1.xlsx");
            wb.write(fout);
        }catch (Exception e){
            e.printStackTrace();
        }
        finally {
            try{
                fout.close();
                wb.close();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }
}

package com.gentech.excel;



import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class second {
    public static void main(String[] args) {
        writeContents();
    }
    public static void writeContents(){
        String[] color={"Red","blue","orange","green","yellow","black","white","violet",
                "gray","silver","tomato","purple","gold","dark blue","pink","dark green",
                "indigo","maroon","brown","light green"};
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try{
            wb=new XSSFWorkbook();
            sh=wb.createSheet("Colors");
            row=sh.createRow(9);
            for(int i=10;i<30;i++){
                cell=row.createCell(i-10);
                cell.setCellValue(color[i-10]);
            }
            fout=new FileOutputStream("D:\\\\excel\\\\second.xlsx");
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

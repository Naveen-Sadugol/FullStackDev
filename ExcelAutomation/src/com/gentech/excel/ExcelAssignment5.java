package com.excel.assignments;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class ExcelAssignment5 {
    public static void main(String[] args) {
        writeContents();
    }
    public static void writeContents(){
        String[] names={"Vijay","Vinay","shreyas","Thilak","suresh","abhishek",
        "sai shashank","revanth","shashidhar","mukesh","suchith","raina","jairam"
        ,"deepak","siraj","likith","srujan","chirag","sham","ram"};
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try{
            wb=new XSSFWorkbook();
            sh=wb.createSheet("names");
            for(int i=10;i<30;i++){
                row=sh.createRow(i-1);
                cell=row.createCell(0);
                cell.setCellValue(names[i-10]);
            }
            fout=new FileOutputStream("E:\\Vijay Kumar A\\ExcelSheets\\Excel5.xlsx");
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

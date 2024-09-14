package com.excel.assignments;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelAssignment8 {
    static String[][] data=new String[20][2];
    public static void main(String[] args) {
        readContents();
        writeContents();
    }
    public static void readContents(){
        FileInputStream fin=null;
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try{
            fin=new FileInputStream("E:\\Vijay Kumar A\\ExcelSheets\\ReadContents\\country.xlsx");
            wb=new XSSFWorkbook(fin);
            sh=wb.getSheet("Sheet1");
            int rc=sh.getPhysicalNumberOfRows();
            for(int i=0;i<rc;i++){
                row=sh.getRow(i);
                int cc=row.getPhysicalNumberOfCells();
                for(int j=0;j<cc;j++){
                    cell=row.getCell(j);
                    data[i][j]=cell.getStringCellValue();
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        finally {
            try{
                fin.close();
                wb.close();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }
    public static void writeContents(){
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try{
            wb=new XSSFWorkbook();
            sh=wb.createSheet("Country");
            for(int i=0;i<2;i++){
                row = sh.createRow(i+3);
                for(int j=0;j<20;j++){
                    cell=row.createCell(j);
                    cell.setCellValue(data[j][i]);
                }
            }
            fout=new FileOutputStream("E:\\Vijay Kumar A\\ExcelSheets\\Excel8.xlsx");
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

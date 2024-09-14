package com.excel.assignments;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelAssignment6 {
    static FileInputStream fin=null;
    static FileOutputStream fout=null;
    static Workbook wb=null;
    static Sheet sh=null;
    static Row row=null;
    static Cell cell=null;
    static Workbook wb1=null;
    static Sheet sh1=null;
    static Row row1=null;
    static Cell cell1=null;
    public static void main(String[] args) {
        readContents();
    }
    public static void readContents(){
        int r=0;
        try{
            fin=new FileInputStream("E:\\Vijay Kumar A\\ExcelSheets\\ReadContents\\Fruits.xlsx");
            wb=new XSSFWorkbook(fin);
            sh=wb.getSheet("Sheet1");
            int rc=sh.getPhysicalNumberOfRows();
            for(int i=0;i<rc;i++){
                row=sh.getRow(i);
                int cc=row.getPhysicalNumberOfCells();
                for(int j=0;j<cc;j++){
                    cell=row.getCell(j);
                    String data=cell.getStringCellValue();
                    writeContent(data,r);
                }
                r++;
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
    public static void writeContent(String data,int r){
        try{
            if(wb1==null){
                wb1=new XSSFWorkbook();
                sh1=wb1.createSheet("Fruits");
                row1=sh1.createRow(r);
                cell1=row1.createCell(0);
                cell1.setCellValue(data);
            }
            else{
                row1=sh1.createRow(r);
                cell1=row1.createCell(0);
                cell1.setCellValue(data);
            }
            fout=new FileOutputStream("E:\\Vijay Kumar A\\ExcelSheets\\Excel6.xlsx");
            wb1.write(fout);
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

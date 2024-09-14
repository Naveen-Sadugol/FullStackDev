package com.excel.assignments;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class ExcelAssignment3 {
    public static void main(String[] args) {
        writeContentsDiogonally();
    }
    public static void writeContentsDiogonally(){
        String[] city={"Bangalore","Mumbai","Delhi","Chennai","Kolkata","Hyderabad",
        "Pune","Jaipur","Ahmedabhad","Lucknow","Kanpur","Patna","Ludhiana","Vadodara",
        "Varnasi","Vishakapatnam","Indore","Bhopal","Agra","Coimbatore"};
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try{
            wb=new XSSFWorkbook();
            sh=wb.createSheet("City");
            for(int i=0;i<20;i++){
                row=sh.createRow(i);
                cell=row.createCell(i);
                cell.setCellValue(city[i]);
            }
            fout=new FileOutputStream("E:\\Vijay Kumar A\\ExcelSheets\\Excel3.xlsx");
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

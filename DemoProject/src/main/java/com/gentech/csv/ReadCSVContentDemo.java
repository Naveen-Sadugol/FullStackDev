package com.gentech.csv;

import com.opencsv.CSVReader;

import java.io.FileReader;

public class ReadCSVContentDemo {
    public static void main(String[] args) {
        readContent();
    }

    private static void readContent()
    {
        FileReader fr=null;
        CSVReader reader=null;
        try
        {
            fr=new FileReader("E:\\Example\\Employees.csv");
            reader=new CSVReader(fr);

            String data[]=null;
            while((data=reader.readNext())!=null)
            {
                for(int i=0;i<data.length;i++)
                {
                    System.out.print(data[i]+",");
                }
                System.out.println();
            }
        }catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            try
            {
                fr.close();
                reader.close();
            }catch (Exception e)
            {
                e.printStackTrace();
            }
        }
    }
}

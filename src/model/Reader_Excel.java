package model;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//SUPPORT XLS & XLSX
public class Reader_Excel
{
    private Workbook wb;
    private Sheet sheet;

    public Workbook Wb() { return wb; }
    public Sheet get_sheet() { return sheet; }

    public Reader_Excel(String path2file, boolean is_xls) throws IOException
    {
        set_workbook(path2file, is_xls);
        set_firstSheet();
    }

    public Reader_Excel(File file, boolean is_xls) throws IOException
    {
        set_workbook(file, is_xls);
        set_firstSheet();
    }

    public Reader_Excel(String path2file, boolean is_xls, int sheet_num) throws IOException
    {
        set_workbook(path2file, is_xls);
        set_sheet(sheet_num - 1);
    }

    public Reader_Excel(File file, boolean is_xls, int sheet_num) throws IOException
    {
        set_workbook(file, is_xls);
        set_sheet(sheet_num - 1);
    }

    public void set_workbook(String path2file, boolean is_xls) throws IOException
    {
//        set_workbook(new FileInputStream(new File(path2file)), is_xls);
        set_workbook(new FileInputStream(new File(path2file)), path2file.substring(path2file.lastIndexOf(".") + 1));
    }

    public void set_workbook(File file, boolean is_xls) throws IOException
    {
//        set_workbook(new FileInputStream(file), is_xls);
        String filename = file.getName();
        set_workbook(new FileInputStream(file), filename.substring(filename.lastIndexOf(".") + 1));
    }

    public void set_workbook(FileInputStream FIS, boolean is_xls) throws IOException
    {
        if(is_xls)
            wb = new HSSFWorkbook(FIS);
        else
            wb = new XSSFWorkbook(FIS);
    }

    public void set_workbook(FileInputStream FIS, String ext) throws IOException
    {
        if(ext.equalsIgnoreCase("xls"))
            wb = new HSSFWorkbook(FIS);
        else if(ext.equalsIgnoreCase("xlsx"))
            wb = new XSSFWorkbook(FIS);
    }

    public void set_firstSheet()
    {
        set_sheet(0);
    }

    public void set_lastSheet()
    {
        set_sheet(wb.getNumberOfSheets());
    }

    public void set_sheet(int sheet_i)
    {
        sheet = wb.getSheetAt(sheet_i);
    }

    public void get_colsName()
    {
        
    }

    public void itr_rowCell()
    {
        Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
        while (itr.hasNext())
        {
            Row row = itr.next();
            Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
            while (cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                //                    cell.setCellType(CellType.STRING);
                DataFormatter dataFormatter = new DataFormatter();
                System.out.print(dataFormatter.formatCellValue(cell) + "\t\t\t");
//                    switch (cell.getCellType())
//                    {
//                        case STRING:    //field that represents string cell type  
//                            System.out.print(cell.getStringCellValue() + "\t\t\t");
//                            break;
//                        case NUMERIC:    //field that represents number cell type  
//                            System.out.print(cell.getNumericCellValue() + "\t\t\t");
//                            break;
//                        case _NONE:
//                            System.out.print(cell.getDateCellValue()+ "\t\t\t");
//                        default:
//                    }
            }
            System.out.println("");
        }
    }

    public void itr_rowCell_3()
    {
        DataFormatter dataFormatter = new DataFormatter();
        int total_rows_i = sheet.getLastRowNum();
        for(int row_i = 0; row_i <= total_rows_i; row_i++)
        {
            Row row_data = sheet.getRow(row_i);
            for(int col_i = 0; col_i < row_data.getLastCellNum(); col_i++)
                System.out.print(dataFormatter.formatCellValue(row_data.getCell(col_i)) + "\t");
            System.out.println("");
        }
    }

    public static void main(String[] args) {
        try {
            Reader_Excel imp_exl = new Reader_Excel("C:\\Users\\ckit\\Documents\\NetBeansProjects\\SubModules\\src\\052DS_AMS20190331.xls", true);
            imp_exl.itr_rowCell_3();
            imp_exl = new Reader_Excel("C:\\Users\\ckit\\Documents\\NetBeansProjects\\SubModules\\src\\Covid-19 Vaccination Record Survey (Relatives).xlsx", false);
            imp_exl.itr_rowCell();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
package extensions;

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
public abstract class Reader_Excel
{
    public DataFormatter dataFormatter = new DataFormatter();

    private String filename;
    private Workbook wb;
    private Sheet sheet;
    private int total_rows;
    protected Row row;

    public String get_reading_filename() { return filename; }
    public Workbook wb() { return wb; }
    public Sheet sheet() { return sheet; }
    public int total_rows() { return total_rows; }

    public Reader_Excel(String path2file) throws IOException
    {
        set_workbook(path2file);
        set_firstSheet();
        set_total_rows();
    }

    public Reader_Excel(File file) throws IOException
    {
        set_workbook(file);
        set_firstSheet();
        set_total_rows();
    }

    public Reader_Excel(String path2file, int sheet_num) throws IOException
    {
        set_workbook(path2file);
        set_sheet(sheet_num - 1);
        set_total_rows();
    }

    public Reader_Excel(File file, int sheet_num) throws IOException
    {
        set_workbook(file);
        set_sheet(sheet_num - 1);
        set_total_rows();
    }

    protected void set_workbook(String path2file) throws IOException
    {
        File file = new File(path2file);
        filename = file.getName();
        set_workbook(new FileInputStream(file), path2file.substring(path2file.lastIndexOf(".") + 1));
    }

    protected void set_workbook(File file) throws IOException
    {
        filename = file.getName();
        set_workbook(new FileInputStream(file), filename.substring(filename.lastIndexOf(".") + 1));
    }

    protected void set_workbook(FileInputStream FIS, String ext) throws IOException
    {
        if(ext.equalsIgnoreCase("xls"))
            wb = new HSSFWorkbook(FIS);
        else if(ext.equalsIgnoreCase("xlsx"))
            wb = new XSSFWorkbook(FIS);
    }

    protected void set_firstSheet()
    {
        set_sheet(0);
    }

    protected void set_lastSheet()
    {
        set_sheet(wb.getNumberOfSheets());
    }

    protected void set_sheet(int sheet_i)
    {
        sheet = wb.getSheetAt(sheet_i);
    }

    protected void set_total_rows()
    {
        total_rows = sheet.getLastRowNum();
    }

    public String get_cellData(int col_i)
    {
        return dataFormatter.formatCellValue(row.getCell(col_i));
    }

    // <editor-fold defaultstate="collapsed" desc="TEST RUN">

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
                dataFormatter = new DataFormatter();
                System.out.print(dataFormatter.formatCellValue(cell) + "\t\t\t");
            }
            System.out.println("");
        }
    }

    public void itr_rowCell_3()
    {
        dataFormatter = new DataFormatter();
        int total_rows_i = sheet.getLastRowNum();
        for(int row_i = 0; row_i <= total_rows_i; row_i++)
        {
            Row row_data = sheet.getRow(row_i);
            for(int col_i = 0; col_i < row_data.getLastCellNum(); col_i++)
                System.out.print(dataFormatter.formatCellValue(row_data.getCell(col_i)) + "\t");
            System.out.println("");
        }
    }

//    public static void main(String[] args)
//    {
//        try {
//            Reader_Excel imp_exl = new Reader_Excel("C:\\Users\\ckit\\Documents\\NetBeansProjects\\SubModules\\Excel\\Excel_Reader\\src\\052DS_AMS20190331.xls");
//            imp_exl.itr_rowCell_3();
//            imp_exl = new Reader_Excel("C:\\Users\\ckit\\Documents\\NetBeansProjects\\SubModules\\Excel\\Excel_Reader\\src\\Covid-19 Vaccination Record Survey (Relatives).xlsx");
//            imp_exl.itr_rowCell();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }

    // </editor-fold>
}
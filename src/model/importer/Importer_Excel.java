package model.importer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Importer_Excel
{
    private XSSFWorkbook wb;
    private XSSFSheet sheet;

    public XSSFWorkbook Wb() { return wb; }
    public XSSFSheet Sheet() { return sheet; }

    public Importer_Excel(String path2file) throws FileNotFoundException, IOException
    {
//        path2file = "C:\\Users\\ckit\\Documents\\NetBeansProjects\\SubModules\\src\\IT Task List & Acitivity Plan.xlsx";
        path2file = "C:\\Users\\ckit\\Documents\\NetBeansProjects\\SubModules\\src\\Covid-19 Vaccination Record Survey (Relatives).xlsx";
        FileInputStream FIS = new FileInputStream(new File(path2file));
        wb = new XSSFWorkbook(FIS);
        set_firstSheet();
    }

    public Importer_Excel(File file) throws FileNotFoundException, IOException
    {
        FileInputStream FIS = new FileInputStream(file);
        wb = new XSSFWorkbook(FIS);
    }

    public Importer_Excel(FileInputStream FIS) throws IOException
    {
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

    public void itr_rowCell_2()
    {
        DataFormatter dataFormatter = new DataFormatter();
        Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
        int total_rows_i = sheet.getLastRowNum();
        for(int row_i = 0; row_i <= total_rows_i; row_i++)
        {
            XSSFRow row_data = sheet.getRow(row_i);
//            System.out.println("ROW: " + (row_i + 1) + " | phyxNumCells: " + row_data.getLastCellNum());
            for(int col_i = 0; col_i < row_data.getLastCellNum(); col_i++)
            {
                System.out.print(dataFormatter.formatCellValue(row_data.getCell(col_i)) + "\t");
            }
            System.out.println("");
        }
    }

    public static void main(String[] args) {
        try {
            Importer_Excel imp_exl = new Importer_Excel("");
            imp_exl.itr_rowCell_2();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
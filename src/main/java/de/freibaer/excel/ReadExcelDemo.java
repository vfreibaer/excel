package de.freibaer.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo 
{
	
	static ArrayList<String[]> arrayList;  
	
    public static void main(String[] args) 
    {
        try
        {
            File file2 = new File("C:/porjekte-luna/excel/excel/src/main/resources/Rohdaten-Etiketten1-7er_gesamt.xls");
			FileInputStream file = new FileInputStream(file2);
 
			Workbook exWorkBook = WorkbookFactory.create(file);
			
            //Create Workbook instance holding reference to .xlsx file
//            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //Get first/desired sheet from the workbook
            Sheet sheet = exWorkBook.getSheetAt(0);
            
            
            arrayList = new ArrayList<String[]>();
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "");
                            
                            String stringCellValue = cell.getStringCellValue();
                            String[] split = stringCellValue.split("-");
                            arrayList.add(split);
                            break;
                    }
                }
                System.out.println("");
            }
            
            
            
            file.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
        
      //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Data Lineup");
          
        //This data needs to be written (Object[])
          
        //Iterate over data and write to sheet
        int indexcounter=0;
        int rowcounter=0;
        int zerosupportcounter = 0;
        int readedCells = arrayList.size();
        int writtenCells = 0;
        
        while (false ==  arrayList.isEmpty()) 
        {
			Row row = sheet.createRow(rowcounter);

			int cellnum = 0;
			for (int j = 1; j < 8; j++) {
				int readCounter;
				String[] element = null;
				if(arrayList.size()>0)
				{
					element = arrayList.get(indexcounter);
					readCounter = Integer.valueOf( element[2] );
				}
				else{
					readCounter = j+1;
				}

				if(j == readCounter)
				{
					Cell cell = row.createCell(cellnum++);
					String string = element[0]+"-" + element[1] + "-"+ element[2];
					System.out.println(string);
					cell.setCellValue(string );
					arrayList.remove(indexcounter);
				}
				else{
					Cell cell = row.createCell(cellnum++);
					cell.setCellValue("00-00-0" );
					zerosupportcounter++;
				}
				writtenCells++;
				
			}
			rowcounter++;
        	            
        	
		}
        
        
        
//        int rownum = 0;
//        Row row = sheet.createRow(rownum);
//        for (int i = 0; i < arrayList.size(); i++) {
//        	String[] strings = arrayList.get(i);
//
//        	int cellnum = 0;
//            for (int j = 0; j < 6; j++) 
//            {
//				
//               Cell cell = row.createCell(cellnum++);
//               if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Integer)
//                    cell.setCellValue((Integer)obj);
//            }
//        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("C:/porjekte-luna/excel/excel/target/FINAL-Etiketten1-7er_gesamt.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("FINAL-Etiketten1-7er_gesamt.xls written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
}
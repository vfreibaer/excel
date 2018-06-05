package de.freibaer.excel;

import java.awt.HeadlessException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfAction;
import com.itextpdf.text.pdf.PdfDestination;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;



public class App extends JFrame
{
    public static void main(String[] args) throws Exception{
        App app = new App();
        app.sayHello("Hello","Choose File!");
        File choosenFile = app.chooseFile();
        app.sayHello("choosenFile","File: " + choosenFile.getName());
        //FileUtils.writeStringToFile(new File("test.txt"), new ExcelToHtml(new FileInputStream(choosenFile)).getHTML());


        FileInputStream filecontent = new FileInputStream(choosenFile);
        
        choosenFile.getPath().replace(".xls", ".pdf");
        FileOutputStream out = new FileOutputStream(new File(choosenFile.getPath().replace(".xls", ".pdf") ));
        HSSFWorkbook my_xls_workbook = null;
        HSSFSheet my_worksheet = null;
        XSSFWorkbook my_xlsx_workbook = null;
        XSSFSheet my_worksheet_xlsx = null;
        Document document = new Document(PageSize.ARCH_E, 0, 0, 0, 0);
        PdfWriter writer = PdfWriter.getInstance(document, out);
        document.open();
        PdfDestination magnify = null;
        float magnifyOpt = (float) 70.0;
        magnify = new PdfDestination(PdfDestination.XYZ, -1, -1, magnifyOpt / 100);
        int pageNumberToOpenTo = 1;
        PdfAction zoom = PdfAction.gotoLocalPage(pageNumberToOpenTo, magnify, writer);
        writer.setOpenAction(zoom);
        document.addDocListener(writer);

        Iterator<Row> rowIterator = null;
        int maxColumn = 0;
        if (choosenFile.getName().contains(".xls")) {
            try {
                my_xls_workbook = new HSSFWorkbook(filecontent);
                my_worksheet = my_xls_workbook.getSheetAt(3);
                rowIterator = my_worksheet.iterator();
                maxColumn = my_worksheet.getRow(0).getLastCellNum();
                maxColumn = 1;
            } catch (IOException ex) {
//                Logger.getLogger(PdfConversion.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

        if (choosenFile.getName().contains(".xlsx")) {
            try {
                my_xlsx_workbook = new XSSFWorkbook(filecontent);
                my_worksheet_xlsx = my_xlsx_workbook.getSheetAt(0);
                rowIterator = my_worksheet_xlsx.iterator();
                maxColumn = my_worksheet_xlsx.getRow(0).getLastCellNum();
            } catch (IOException ex) {
//                Logger.getLogger(PdfConversion.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        PdfPTable my_table = new PdfPTable(maxColumn);
        my_table.setHorizontalAlignment(Element.ALIGN_CENTER);
        my_table.setWidthPercentage(100);
        my_table.setSpacingBefore(0f);
        my_table.setSpacingAfter(0f);
        PdfPCell table_cell;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next(); //Fetch CELL
                switch (cell.getCellType()) { //Identify CELL type
                    case Cell.CELL_TYPE_STRING:
                        //Push the data from Excel to PDF Cell
                        table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                        if (row.getRowNum() == 0) {
                            table_cell.setBackgroundColor(BaseColor.LIGHT_GRAY);
                            table_cell.setBorderColor(BaseColor.BLACK);
                        }
                        my_table.addCell(table_cell);
                        break;
                }
            }
        }
        document.add(my_table);
        document.close();
        System.out.println("Excel file converted to PDF successfully");
        
    }

    public App() throws HeadlessException {
        setVisible(false);
    }
    
    private File chooseFile() {
    	JSystemFileChooser chooser = new JSystemFileChooser();
	    FileNameExtensionFilter filter = new FileNameExtensionFilter(
	        "Excel File", "xls", "xlsx");
	    chooser.setFileFilter(filter);
	    int returnVal = chooser.showOpenDialog(this);
	    if(returnVal == JFileChooser.APPROVE_OPTION) {
	       System.out.println("You chose to open this file: " +chooser.getSelectedFile().getName());
	       return chooser.getSelectedFile();
	    }
	    return null;
    }

    private void sayHello(String title, String message) {
        JOptionPane.showMessageDialog(this, message, title, JOptionPane.INFORMATION_MESSAGE);
        dispose();
    }

}

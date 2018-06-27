package de.freibaer.excel;

import java.awt.HeadlessException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.poiji.bind.Poiji;
import com.poiji.exception.PoijiExcelType;
import com.poiji.option.PoijiOptions;
import com.poiji.option.PoijiOptions.PoijiOptionsBuilder;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;



public class App extends JFrame
{
    public static void main(String[] args) throws Exception{
        App app = new App();
        
        File choosenFile = new File("C:/Users/Vitus/Desktop/Regenwasseranlage-Stand_05.2018.xls");
        if(!choosenFile.exists()) {
        	app.sayHello("Hello","Choose File!");
        	choosenFile = app.chooseFile();
            app.sayHello("choosenFile","File: " + choosenFile.getName());
        }
        
        List<PoijRow> extractRows = app.extractRows(choosenFile);
        app.sayHello("Zeilen ausgelesen","Read lines: " + extractRows.size());
        String path = choosenFile.getParent() + File.separator +"Regenwasseranlage_Report_" +System.currentTimeMillis();
        app.sayHello("Erzeuge Ordner","Erzeuge Ordner: " + path );
        File report = new File(path);
        report.mkdir();
        
//        File template = new File(app.getClass().getClassLoader().getResource("Template.xls").getPath());
        
        for (int i = 1; i < extractRows.size(); i++) {
			File writeTemplate = new File(report.getPath() + File.separator + "Wartungsprotokoll_Sickergrube_" +i+".xls");
//        	Files.copy(template.toPath(), writeTemplate.toPath());
        	
        	POIFSFileSystem fs = new POIFSFileSystem(app.getClass().getClassLoader().getResourceAsStream("Template.xls"));
        	HSSFWorkbook wb = new  HSSFWorkbook(fs, true);
        	//Will load an xls, preserving its structure (macros included). You can then modify it,
        	HSSFSheet sheet1 = wb.getSheet("Wartungsprotokoll Sickergrube");
		    //...
        	PoijRow poijRow = extractRows.get(i);
        	sheet1.getRow(0).getCell(1).setCellValue(poijRow.getColB());
        	sheet1.getRow(1).getCell(1).setCellValue(poijRow.getColC());
        	sheet1.getRow(2).getCell(1).setCellValue(poijRow.getColD());
        	sheet1.getRow(3).getCell(1).setCellValue(poijRow.getColE());
        	sheet1.getRow(4).getCell(1).setCellValue(poijRow.getColF());
        	sheet1.getRow(5).getCell(1).setCellValue(poijRow.getColG());
        	sheet1.getRow(6).getCell(1).setCellValue(poijRow.getColH());
        	sheet1.getRow(7).getCell(1).setCellValue(poijRow.getColI());
        	sheet1.getRow(8).getCell(1).setCellValue(poijRow.getColJ());
        	sheet1.getRow(9).getCell(1).setCellValue(poijRow.getColK());
        	
//        	colL=, colM=, colN=, colO=, colP=, colQ=09/2017, colR=, colS=, colT=, colU=, colV=, colW=]
        	String strLastCleanup = "";
        	strLastCleanup =!poijRow.getColL().isEmpty() ?poijRow.getColL():strLastCleanup;
        	strLastCleanup =!poijRow.getColM().isEmpty() ?poijRow.getColM():strLastCleanup;
        	strLastCleanup =!poijRow.getColN().isEmpty() ?poijRow.getColN():strLastCleanup;
        	strLastCleanup =!poijRow.getColO().isEmpty() ?poijRow.getColO():strLastCleanup;
        	strLastCleanup =!poijRow.getColP().isEmpty() ?poijRow.getColP():strLastCleanup;
        	strLastCleanup =!poijRow.getColQ().isEmpty() ?poijRow.getColQ():strLastCleanup;
        	strLastCleanup =!poijRow.getColR().isEmpty() ?poijRow.getColR():strLastCleanup;
        	strLastCleanup =!poijRow.getColS().isEmpty() ?poijRow.getColS():strLastCleanup;
        	strLastCleanup =!poijRow.getColT().isEmpty() ?poijRow.getColT():strLastCleanup;
        	strLastCleanup =!poijRow.getColU().isEmpty() ?poijRow.getColU():strLastCleanup;
        	strLastCleanup =!poijRow.getColV().isEmpty() ?poijRow.getColV():strLastCleanup;
        	strLastCleanup =!poijRow.getColW().isEmpty() ?poijRow.getColW():strLastCleanup;
        	
        	sheet1.getRow(10).getCell(1).setCellValue(strLastCleanup);
//        	sheet1.getRow(11).getCell(1).setCellValue(poijRow.getColC());
		
		    FileOutputStream fileOut = new FileOutputStream(writeTemplate);
		    wb.write(fileOut);
		    fs.close();
		    fileOut.close();
        	
        	
		}
    }
    
	private List<PoijRow> extractRows(File choosenFile ) {
//		Log.jlog.debug("Extrahiere Zeilen aus Excel-Datei: " +  fullFileName);
		
		final PoijiOptions options = PoijiOptionsBuilder.settings().preferNullOverDefault(false).sheetIndex(1).build();
		
		final List<PoijRow> rows = new ArrayList<>();
    	try (InputStream inStream = new FileInputStream(choosenFile)) {
    		rows.addAll(Poiji.fromExcel(inStream, PoijiExcelType.XLS, PoijRow.class, options));
    	}
    	catch (final Exception e) {
    		System.out.println("Exception while extracting rows: ");
    	}
    	
    	return Collections.unmodifiableList(rows);
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

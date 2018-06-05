package de.freibaer.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;


public class App 
{
    public static void main( String[] args )
    {
    	try {
			FileInputStream aaaa = new FileInputStream("Rohdaten-Etiketten1-7er_gesamt.xls");
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
    
    }
}

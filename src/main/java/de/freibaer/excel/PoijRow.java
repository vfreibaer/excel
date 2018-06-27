package de.freibaer.excel;


import com.poiji.annotation.ExcelCell;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public final class PoijRow extends AbstractRow {
	// Col_A: Vertriebsname
	@ExcelCell(0)
	private String colA;
	// Col_B: Hauptversicherung
	@ExcelCell(1)
	private String colB;
	// Col_C: Tarif HV
	@ExcelCell(2)
	private String colC;
	// Col_D: Fondsklasse
	@ExcelCell(3)
	private String colD;
	// Col_E: Produktmerkmale
	@ExcelCell(4)
	private String colE;
	// Col_F: Zusatzversicherung
	@ExcelCell(5)
	private String colF;
	// Col_G: Tarif NV
	@ExcelCell(6)
	private String colG;
	// Col_H: Abschluss-Alter HV min.
	@ExcelCell(7)
	private String colH;
	// Col_I: Risikoklasse, Laufzeit ab 12 Jahre
	@ExcelCell(8)
	private String colI;
	// Col_J: Risikoklasse, Laufzeit unter 12 Jahre
	@ExcelCell(9)
	private String colJ;
	// Col_K: Absicherung des Einkommensverlustes: Sehr wichtig, wichtig, neutral
	@ExcelCell(10)
	private String colK;
	// Col_L: regelmaesige Beitragszahlung
	@ExcelCell(11)
	private String colL;
	// Col_M: Einmalzahlung
	@ExcelCell(12)
	private String colM;
	@ExcelCell(13)
	private String colN;
	@ExcelCell(14)
	private String colO;
	@ExcelCell(15)
	private String colP;
	@ExcelCell(16)
	private String colQ;
	@ExcelCell(17)
	private String colR;
	@ExcelCell(18)
	private String colS;
	@ExcelCell(19)
	private String colT;
	@ExcelCell(20)
	private String colU;
	@ExcelCell(21)
	private String colV;
	@ExcelCell(22)
	private String colW;
	
	
	
	
	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("\nAltersvorsorge [rowIndex=").append(rowIndex)
			   .append(", colA=").append(colA)
			   .append(", colB=").append(colB)
			   .append(", colC=").append(colC)
			   .append(", colD=").append(colD)
			   .append(", colE=").append(colE)
			   .append(", colF=").append(colF)
			   .append(", colG=").append(colG)
			   .append(", colH=").append(colH)
			   .append(", colI=").append(colI)
			   .append(", colJ=").append(colJ)
			   .append(", colK=").append(colK)
			   .append(", colL=").append(colL)
			   .append(", colM=").append(colM)
			   .append(", colN=").append(colN)
			   .append(", colO=").append(colO)
			   .append(", colP=").append(colP)
			   .append(", colQ=").append(colQ)
			   .append(", colR=").append(colR)
			   .append(", colS=").append(colS)
			   .append(", colT=").append(colR)
			   .append(", colU=").append(colU)
			   .append(", colV=").append(colV)
			   .append(", colW=").append(colW)
			   .append("]\n");
		
		return builder.toString();
	}
		
}

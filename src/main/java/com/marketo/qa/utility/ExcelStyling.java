package com.marketo.qa.utility;


import java.io.FileInputStream;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;

public class ExcelStyling {
	
	
	static String ExcelPath = System.getProperty("user.dir") + "//Reports//";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());

	XSSFRow myRow = null;
<<<<<<< HEAD
	
	//writing Excel data with color code
=======
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void WriteExcel(Workbook workbook,XSSFSheet sheet ,int row, int col, String colorCode, String value,int height) throws DecoderException ,Exception {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		Font headerFont = workbook.createFont();
		byte[] rgbB = Hex.decodeHex(colorCode); // get byte array from hex string
        XSSFColor color = new XSSFColor(rgbB, null);
          
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(color);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) height);
	    XSSFRow myRow = sheet.getRow(row);
	    
	    if (myRow == null)
	        myRow = sheet.createRow(row);
	    
	    XSSFCell myCell = myRow.createCell(col);
	    cellStyle.setFont(headerFont);
	    myCell.setCellStyle(cellStyle);
	    myCell.setCellValue(value);
	}

<<<<<<< HEAD
	//writing Excel data with out color code
=======
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void WriteExcel(Workbook workbook,XSSFSheet sheet ,int row, int col, int value) throws Exception {
		
		XSSFRow	myRow=null;
	     myRow = sheet.getRow(row);

	    if (myRow == null)
	        myRow = sheet.createRow(row);
	    

	    XSSFCell myCell = myRow.createCell(col);
	    myCell.setCellValue(value);
	}

<<<<<<< HEAD
	//Deleting column method
=======

>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void deleteColumn(XSSFSheet sheet, int columnToDelete) {
		for (int rId = 0; rId < sheet.getLastRowNum(); rId++) {
	        Row row = sheet.getRow(rId);
	        for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
	            Cell cOld = row.getCell(cID);
	            if (cOld != null) {
	                row.removeCell(cOld);
	            }
	            Cell cNext = row.getCell(cID + 1);
	            if (cNext != null) {
	                Cell cNew = row.createCell(cID, cNext.getCellType());
	                cloneCell(cNew, cNext);
	                //Set the column width only on the first row.
	                //Other wise the second row will overwrite the original column width set previously.
	                if(rId == 0) {
	                    sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));

	                }
	            }
	        }}
	        }
	    
<<<<<<< HEAD
   //Cloning cell from row
=======
	



>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void cloneCell(Cell cNew, Cell cOld) {
	    cNew.setCellComment(cOld.getCellComment());
	    cNew.setCellStyle(cOld.getCellStyle());

	    if (CellType.BOOLEAN == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getBooleanCellValue());
	    } else if (CellType.NUMERIC == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getNumericCellValue());
	    } else if (CellType.STRING == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getStringCellValue());
	    } else if (CellType.ERROR == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getErrorCellValue());
	    } else if (CellType.FORMULA == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getCellFormula());
	    }
	}
<<<<<<< HEAD
        //Defining cell style for pasting data
=======

	        
	
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static CellStyle createColorWithtext(Workbook workbook,int rownum,int cellno,XSSFSheet sheet,String text,String rgbS) throws DecoderException , Exception {
	      //  CellStyle style = workbook.createCellStyle();
	        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
	        
	        Row row = sheet.createRow(rownum); 
	        
	        
	        byte[] rgbB = Hex.decodeHex(rgbS); // get byte array from hex string
	        XSSFColor color = new XSSFColor(rgbB, null);
	     
           Cell cell = row.createCell(cellno);  
            cell.setCellValue(text);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	        cellStyle.setFillForegroundColor(color);
	        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);  
	        return cellStyle;
	    
	}
<<<<<<< HEAD
	        
	//Cell merging methods for heading
=======
	    
	    

	public static void setMerge(Workbook workbook,XSSFSheet sheet, int numRow, int untilRow, int numCol, int untilCol) throws Exception {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		 
		XSSFRow row = sheet.createRow(2);
//		cellStyle.setAlignment(HorizontalAlignment.CENTER);
//        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
	    CellRangeAddress cellMerge = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
	    XSSFCell cell = row.createCell(2);
	    cell.setCellStyle(cellStyle);
	    sheet.addMergedRegion(cellMerge);
	    
	   
	 
 
	}

	
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void mergeAndCenter(Workbook workbook,XSSFSheet sheet,String rgbS ,String text, int firstRow, int lastRow, int firstCol, int lastCol,boolean value,int fontHeight) throws DecoderException ,Exception{
	   
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		XSSFFont font = (XSSFFont) workbook.createFont();
        Row row = sheet.createRow(firstRow); 
        
        CellRangeAddress cellRangeAddress = new CellRangeAddress( firstRow,lastRow,firstCol,lastCol);
        sheet.addMergedRegionUnsafe(cellRangeAddress);
        byte[] rgbB = Hex.decodeHex(rgbS); // get byte array from hex string
        XSSFColor color = new XSSFColor(rgbB, null);
       
     
        font.setBold(value);
        font.setFontHeightInPoints((short) fontHeight);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(color); 
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      

      //  cell.setCellStyle(cellStyle);
        
        // Creates the cell
        Cell cell1 = CellUtil.createCell(row, 2, text, cellStyle);
        
        // Sets the allignment to the created cell
        CellUtil.setVerticalAlignment(cell1, VerticalAlignment.CENTER);
       // return cellStyle;
    
	}

<<<<<<< HEAD
	//pasting images 
=======
	
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void pestImg(Workbook workbook,XSSFSheet sheet,int col,int row,String path,double scale) throws IOException {
		
		//add picture data to this workbook.
		InputStream is = new FileInputStream(path);
		byte[] bytes = IOUtils.toByteArray(is);
		int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		is.close();
		
		CreationHelper helper = workbook.getCreationHelper();
		// Create the drawing patriarch. This is the top level container for all shapes. 
		Drawing drawing = sheet.createDrawingPatriarch();

		//add a picture shape
		ClientAnchor anchor = helper.createClientAnchor();
		//set top-left corner of the picture,
		//subsequent call of Picture#resize() will operate relative to it
		anchor.setCol1(col);
		anchor.setRow1(row);

		
		Picture pict = drawing.createPicture(anchor, pictureIdx);
		scale = 0.6;
		//Reset the image to the original size
		pict.resize();
		if (scale < 1) {
		    pict.resize(scale);
		}
		
		
		
	}
<<<<<<< HEAD
	 
=======
	
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void WriteExcel(Workbook workbook,XSSFSheet sheet ,int row, int col,String value,int height) throws Exception  {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		Font headerFont = workbook.createFont();
		 
          
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
       cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
        headerFont.setFontHeightInPoints((short) height);
        XSSFRow	myRow=null;
	     myRow = sheet.getRow(row);
	     
	    if (myRow == null)
	        myRow = sheet.createRow(row);
	    	
	    XSSFCell myCell = myRow.createCell(col);
	    		
	    cellStyle.setFont(headerFont);
	    myCell.setCellStyle(cellStyle);
	    myCell.setCellValue(value);
	}
	
<<<<<<< HEAD
=======

>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static void WriteExcel(Workbook workbook,XSSFSheet sheet ,int row, int col,int value,int height) throws Exception  {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		Font headerFont = workbook.createFont();
		
          
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
       
        headerFont.setFontHeightInPoints((short) height);
        XSSFRow	myRow=null;
	     myRow = sheet.getRow(row);
	     
	    if (myRow == null)
	        myRow = sheet.createRow(row);
	    	
	    XSSFCell myCell = myRow.createCell(col);
	    		
	    cellStyle.setFont(headerFont);
	    myCell.setCellStyle(cellStyle);
	     
	    myCell.setCellValue(value);
	}
	
<<<<<<< HEAD
	public static void WriteExcel(Workbook workbook,XSSFSheet sheet ,int row, int col,float value,int height) throws Exception  {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		Font headerFont = workbook.createFont();
		
          
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
       
        headerFont.setFontHeightInPoints((short) height);
        XSSFRow	myRow=null;
	     myRow = sheet.getRow(row);
	     
	    if (myRow == null)
	        myRow = sheet.createRow(row);
	    	
	    XSSFCell myCell = myRow.createCell(col);
	    		
	    cellStyle.setFont(headerFont);
	    myCell.setCellStyle(cellStyle);
	     
	    myCell.setCellValue(value);
	}
	
	public static void WriteExcelWithFormula(Workbook workbook,XSSFSheet sheet ,int row, int col,int height)  {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		Font formulafont = workbook.createFont();
		
=======
	public static void WriteExcelWithFormula(Workbook workbook,XSSFSheet sheet ,int row, int col,int height,String formula)  {
		XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
		Font headerFont = workbook.createFont();
		
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
           
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        
<<<<<<< HEAD
        formulafont.setFontHeightInPoints((short) height);
=======
        headerFont.setFontHeightInPoints((short) height);
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	    XSSFRow myRow = sheet.getRow(row);
	     
	    if (myRow == null)
	        myRow = sheet.createRow(row);
<<<<<<< HEAD
	    XSSFCell myCell = myRow.createCell(col);
	  
	   
	   // myCell.setCellFormula();	
	    
	    cellStyle.setDataFormat(workbook.createDataFormat().getFormat("00.00%"));
	    
	    cellStyle.setFont(formulafont); 
	    myCell.setCellStyle(cellStyle);
	    workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
	    //myCell.setCellValue( myCell.setCellFormula(formula));
=======
	    
	   
	   XSSFCell myCell = myRow.createCell(col);
	    myCell.setCellFormula(formula);	
	    
	    cellStyle.setFont(headerFont); 
	    myCell.setCellStyle(cellStyle);
	    workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
	    //myCell.setCellValue(value);
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	}
	
	
	
	
}

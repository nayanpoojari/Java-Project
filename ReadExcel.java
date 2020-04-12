import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.StringReader;
import java.rmi.server.ExportException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcel {

	public static void main(String[] args){
		readFromExcel();
		//writeFromExcel();
	}
	
	    
	public static void readFromExcel(){
		ArrayList<Integer> serial = new ArrayList<>();
		ArrayList<String> headers = new ArrayList<>();
		
		JFileChooser openFilechooser = new JFileChooser();
		openFilechooser.setDialogTitle("Open File");
		openFilechooser.removeChoosableFileFilter(openFilechooser.getFileFilter());
		FileFilter filter = new FileNameExtensionFilter("Excel File (.xlsx)","xlsx");
		
		if(openFilechooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION){
			File inputFile = openFilechooser.getSelectedFile();
			
			try(FileInputStream in = new FileInputStream(inputFile)){
				XSSFWorkbook importedFile = new XSSFWorkbook(in);
				XSSFSheet sheet = importedFile.getSheetAt(15);
				Iterator<Row> rowIterator = sheet.iterator();
				while(rowIterator.hasNext()){
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					Cell cell = cellIterator.next();
					if(row.getRowNum()==0){
						headers.add(cell.getStringCellValue());
					}
					else{
						/*if(cell.getColumnIndex()==0){
							long time = cell.getDateCellValue().getTime();
							for(long t=time;t<10;t++){
						serial.add((int) cell.getNumericCellValue());*/
						}
				}
		    in.close();
		
			System.out.println("Excel file is read successfully");
			System.out.print("List of serial no is : "+serial);
			System.out.println("\n\n");
					
	
	
	}
	catch(IOException ex){
		Logger.getLogger(ReadExcel.class.getName()).log(Level.SEVERE,null,ex);
	}
	}
}
	public static void writeFromExcel(){

        int[] serial = new int[10];
		/*for(int i=0;i<=serial.length;i++)
		{
			//serial[i]=i+1;
		}*/
		XSSFWorkbook wb= new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Results");
		XSSFRow row1;
	
		row1 = sheet.createRow(0);
		for(int i=0;i<serial.length;i++)
		{
			row1 = sheet.createRow(i+1);
			for(int j=0;j<=1;j++)
			{
				Cell cell = row1.createCell(j);
				if(cell.getColumnIndex()==0)
				{
					cell.setCellValue(serial[i]);
				}
			}
		}
		JFileChooser saveFile = new JFileChooser();
		saveFile.setDialogTitle("Save File");
		saveFile.setSelectedFile(new File("Result.xlsx"));
		File output = saveFile.getSelectedFile();
		try(FileOutputStream out = new FileOutputStream(output)){
			wb.write(out);
			out.close();
			}
		catch(FileNotFoundException ex){
			Logger.getLogger(ExportException.class.getName()).log(Level.SEVERE,null,ex);
			}catch(IOException ex){
				Logger.getLogger(ExportException.class.getName()).log(Level.SEVERE,null,ex);
			}

	}
}
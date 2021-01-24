package org.mailretrieval.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmailToExcel {

	public void retrieveData(Folder inbox) {
		
		try {
			
			inbox.open(Folder.READ_ONLY);
	
		
		// retrieve the messages from the folder in an array and print it
		 Message[] messages = inbox.getMessages();
		
		 //Creating an Excel sheet and storing the data.
		 XSSFWorkbook workbook = new XSSFWorkbook();
		 XSSFSheet sheet = workbook.createSheet("Inbox E-mails");
		 
		
		 //Creating Headings
		 Row rowHeading = sheet.createRow(0);
		 rowHeading.createCell(0).setCellValue("Email-From");
		 rowHeading.createCell(1).setCellValue("Subject");
		 rowHeading.createCell(2).setCellValue("Sent Date");
		 rowHeading.createCell(3).setCellValue("Recieved Date");
		
		 for(int i=0; i<4; i++) {
			 CellStyle stylerowHeading = workbook.createCellStyle();
			 Font font = workbook.createFont();
			 font.setBold(true);
			 font.setFontName(XSSFFont.DEFAULT_FONT_NAME);
			 stylerowHeading.setFont(font);
			 stylerowHeading.setAlignment(HorizontalAlignment.CENTER);
			 
			 rowHeading.getCell(i).setCellStyle(stylerowHeading);
			 sheet.autoSizeColumn(i);
			 
			 
		 }
		 
		
		  int r= 1;
		
		  System.out.println("Writing into the Excel file!!");
		  
		  for (int i = 0, n = messages.length; i < n; i++) { 
			
		  Message m = messages[i];
		  Row row = sheet.createRow(r);
		  
		  Cell cellName = row.createCell(0);
		  cellName.setCellValue(m.getFrom()[0].toString());
		  
		  Cell cellSub = row.createCell(1);
		  cellSub.setCellValue(m.getSubject());
		  
		  Cell cellSentDate = row.createCell(2);
		  cellSentDate.setCellValue(m.getSentDate().toString());
		  
		  Cell cellRecDate = row.createCell(3);
		  cellRecDate.setCellValue(m.getReceivedDate().toString());
		  System.out.println("Processing ############################" );
		 
		  r++; 
		  }
		 
		 
		 
		 
		 //Save to Excel file - add your own specific address 
		  FileOutputStream outputfile = new FileOutputStream(new File("C:\\Users\\sanja\\Desktop\\Retrievedata.xlsx"));
		  workbook.write(outputfile);
		 
		
		 outputfile.close();
		 workbook.close();
		 System.out.println("\n Excel file is written succesfully!");
		
	    

	} catch (MessagingException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
		catch(FileNotFoundException fe) {
			fe.printStackTrace();
		}
		catch(IOException e) {
			e.printStackTrace();
		}
}
}

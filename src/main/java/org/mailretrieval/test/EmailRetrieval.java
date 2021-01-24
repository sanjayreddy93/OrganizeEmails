package org.mailretrieval.test;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;
import javax.mail.Store;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


class MyClass implements Runnable{

	public void run() {
		// TODO Auto-generated method stub
		
	}
	
}

public class EmailRetrieval{

	public static void main(String[] args) {
		
		System.out.println("Starting Email to excel class!!");
		
		String host = "imap.gmail.com";
		String mailStoreType = "imap";
		String mailId ="****************@gmail.com";
		String pwd = "******";
		
		FetchingData Er = new FetchingData();
		Er.createConnection(host, mailStoreType,mailId,pwd);
	}

		
		

}

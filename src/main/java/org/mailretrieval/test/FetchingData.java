package org.mailretrieval.test;

import java.util.Properties;

import javax.mail.Folder;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;
import javax.mail.Store;


public class FetchingData {

	public void createConnection(String host,String mailStoragetype,String user,String pwd) {
		try {
		Properties properties = new Properties(); 
		
		//Creating properties filed using Key-Value pairs
		properties.put("mail.imap.host",host);
		properties.put("mail.imap.port","993");
		properties.put("mail.imap.starttls.enable", "true");
		
		Session session = Session.getDefaultInstance(properties);
		
			//Creating imap store object and connecting with IMAP server
			Store store = session.getStore("imaps");
			store.connect(host, user, pwd);
			
			//Creating folder object and opening it
			Folder emailFolder = store.getFolder("INBOX");
			
			//Accesing data from Inbox and rewriting them in Excel
			EmailToExcel Er = new EmailToExcel();
			Er.retrieveData(emailFolder);

		    //close the store and folder objects
		      emailFolder.close(false);
			
		      store.close();
		
		} catch (NoSuchProviderException e) {
	         e.printStackTrace();
	      } catch (MessagingException e) {
	         e.printStackTrace();
	      } catch (Exception e) {
	         e.printStackTrace();
	      }
	}
}

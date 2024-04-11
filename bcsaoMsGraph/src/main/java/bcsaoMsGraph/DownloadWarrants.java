package bcsaoMsGraph;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.ArrayList;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.microsoft.graph.models.Attachment;
import com.microsoft.graph.models.MailFolder;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.requests.AttachmentCollectionPage;
import com.microsoft.graph.requests.MailFolderCollectionPage;
import com.microsoft.graph.requests.MessageCollectionPage;


public class DownloadWarrants {
	public static void initializeGraph(Properties properties) {
		try {
			Graph.initalizeGraphforApp(properties);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static String getInsertQuery (ArrayList<ArrayList<String>> insertedRows) {
		Logger logger = LoggerFactory.getLogger("EmailsLogger");
		
		String begin = "insert into _bcsao_WarrantControl (WarrantControlNumber, ASA_Email, SequenceNumber, Yes_No) values ";
		ArrayList<String> tempRow;
		String temp = "";
		
		for(int i = 0; i < insertedRows.size(); i++) {
			tempRow = insertedRows.get(i);
			temp = "";
			
			if(i > 0)
				temp += ", ";
			
			temp += "(";
			temp += "'" + tempRow.get(0) + "',";
			temp += "'" + tempRow.get(1) + "',";
			temp += "'" + tempRow.get(2) + "',";
			if(tempRow.get(3).equals("Y")) {
				temp += "1";
			} else {
				temp += "0";
			}
			
			temp += ")";
			
			begin += temp;
			
		}
		
		logger.info(begin);
		
		return begin;
	}
	

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// TODO Auto-generated method stub
		Logger logger = LoggerFactory.getLogger("EmailsLogger");
		
		Properties oAuth = new Properties();
		try {
			oAuth.load(DownloadWarrants.class.getResourceAsStream("/oAuth.properties"));
		} catch (IOException e) {
			e.printStackTrace();
		}
		initializeGraph(oAuth);
		logger.info(oAuth.toString());
		
		
		//String fileName = "";
		// Creating the download directory - if not existing.
		File downloadDirectory = new File("C:\\DownloadedWarrants");
		if(!downloadDirectory.exists())
			downloadDirectory.mkdir();
		
		
		try {
			
			/*
			MessageCollectionPage page = Graph.listMailFromDistributionList("saowarrants@stattorney.org");
			//logger.info(Integer.toString(page.getCurrentPage().size() ));
			for(Message m : page.getCurrentPage()) {
				logger.info(m.subject + "\t" + m.from.emailAddress.address);
				
				AttachmentCollectionPage attachments = Graph.getMailAttachments("saowarrants@stattorney.org", m.id);
				for(Attachment a: attachments.getCurrentPage()) {
					fileName = "";
					
					fileName = a.name;
					logger.info("filename: \t" + fileName);
					
					Graph.downloadAttachment_test2("saowarrants@stattorney.org", m.id, a, downloadDirectory, fileName);
				}
				
				
			}
			*/
			
			String mailboxAddress = "saowarrants@stattorney.org";
			String ASAEmail = "", warrantControlNumber = "", sequenceNumber = "";
			String yesNo = "";
			String[] subjectTokens, attachmentTokens;
			String mailboxFolderId = null;
			
			int len;
			
			ArrayList<String> sqlRow;
			ArrayList<ArrayList<String>> insertedRows = new ArrayList<>();;
			
			//Getting the Outlook MailFolderId for the ProcessedWarrants mail folder.
			MailFolderCollectionPage folders =  Graph.getFolders(mailboxAddress);
			for(MailFolder f: folders.getCurrentPage()) {
				//logger.info(f.displayName + "\t" + f.id);
				if(f.displayName.equals("Processed Warrants")) {
					mailboxFolderId = f.id;
					logger.info("MailFolder Info: " + f.displayName + "\t" + mailboxFolderId);
				}
				
			}
						
			
			MessageCollectionPage searchMailBox = Graph.listMailFromDistributionList(mailboxAddress);
			
			//IF Inbox folder has no new mails, then stops execution.
			if(searchMailBox.getCurrentPage().size() == 0) {
				logger.info("No new emails! Ending execution...");
				return;
			}
			
			for (Message m: searchMailBox.getCurrentPage()) {
				
				//Resetting the sqlRow for each email.
				sqlRow = new ArrayList<>();

				ASAEmail = "";
				warrantControlNumber = "";
				sequenceNumber = "";
				yesNo = "";
				
				subjectTokens = m.subject.split(" ");
				//for(int i = 0 ; i < subjectTokens.length; i++) {
					//logger.info(subjectTokens[i]);
				//}
				warrantControlNumber = subjectTokens[0];
				sequenceNumber = subjectTokens[1];
				ASAEmail = m.sender.emailAddress.address;
				
				if(subjectTokens[2].toLowerCase().equals("yes")) {
					yesNo = "Y";
				} else {
					yesNo = "N";
				}
				
				AttachmentCollectionPage attachments = Graph.getMailAttachments(mailboxAddress, m.id);
				for(Attachment a: attachments.getCurrentPage()) {
					
					attachmentTokens = a.name.split("\\.");
					
					len = attachmentTokens.length;
					
					//Checking the attachment file extension to filter out the email signature images.
					if(attachmentTokens[len-1].equals("png")) {
						logger.info("Attachment is an image. Ignoring ...");
						continue;
					}
					logger.info("Downloading email attachment: " + a.name);
					Graph.downloadAttachment_test2(mailboxAddress, m.id, a, downloadDirectory, a.name);
					
					sqlRow.add(warrantControlNumber);
					sqlRow.add(ASAEmail);
					sqlRow.add(sequenceNumber);
					sqlRow.add(yesNo);
					
					//The DB is updated only if the email has a valid non-image attachment.
					insertedRows.add(sqlRow);
				}
				
				
				Graph.moveEmail(mailboxAddress, m.id, mailboxFolderId);
				
			}
//			System.out.println(insertedRows);
			
			String insertQuery = getInsertQuery(insertedRows);
			
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			String connectionURL = "jdbc:sqlserver://SAOJD1;"
					+ "databaseName=PD51_Data;"
					+ "integratedSecurity=true";
			Connection conn = DriverManager.getConnection(connectionURL);
			PreparedStatement ps;
			
			ps = conn.prepareStatement(insertQuery);
			ps.execute();
			logger.info("updated _bcsao_WarrantControl table with " + insertedRows.size() + " rows.");
			
			
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}

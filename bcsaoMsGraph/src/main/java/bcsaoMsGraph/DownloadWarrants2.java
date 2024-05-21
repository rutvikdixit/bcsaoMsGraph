package bcsaoMsGraph;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.microsoft.graph.models.Attachment;
import com.microsoft.graph.models.FileAttachment;
import com.microsoft.graph.models.MailFolder;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.requests.AttachmentCollectionPage;
import com.microsoft.graph.requests.MailFolderCollectionPage;
import com.microsoft.graph.requests.MessageCollectionPage;

/*
 * https://github.com/rutvikdixit/bcsaoMsGraph/
 */
public class DownloadWarrants2 {

	public static void initializeGraph(Properties properties) {
		try {
			Graph.initalizeGraphforApp(properties);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static String getInsertQuery (ArrayList<ArrayList<String>> insertedRows) {
		Logger logger = LoggerFactory.getLogger("InsertQueryFunctionLogger");
		
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
		
		Logger logger = LoggerFactory.getLogger("WarrantsLogger");

		Properties oAuth = new Properties();
		try {
			oAuth.load(DownloadWarrants.class.getResourceAsStream("/oAuth.properties"));
		} catch (IOException e) {
			e.printStackTrace();
		}
		initializeGraph(oAuth);
		logger.info(oAuth.toString());
		
		File downloadDirectory = new File("C:\\DownloadedWarrants");
		if(!downloadDirectory.exists()) {
			downloadDirectory.mkdir();
			logger.info("Download Directory has been created.");
		} else {
			logger.info("Download Directory already exists.");
		}
		
		
		// Declarations
		String mailboxAddress = "saowarrants@stattorney.org";
		String mailSubject;
		String ASAEmail = "", warrantControlNumber = "", sequenceNumber = "";
		String yesNo = "";
		String[] subjectTokens, attachmentTokens;
		String mailboxFolderId = null;
		
		int len, validAttachments;
		boolean fwd, noWarrantFlag;
		
		FileAttachment fileAttachment;
		
		ArrayList<String> sqlRow;
		ArrayList<ArrayList<String>> insertedRows = new ArrayList<>();;
		
		//Pattern to identify Forwards in emails.
		Pattern fwdPattern =  Pattern.compile("fw[a-zA-Z0-9]*:", Pattern.CASE_INSENSITIVE);
		Matcher fwdMatcher;
		
		
		//Starting Main Script
		
		//Getting Processed Warrants FolderID from O365
		MailFolderCollectionPage folders = null;
		try {
			folders = Graph.getFolders(mailboxAddress);
		} catch (Exception e) {
			logger.error("Error fetching mailbox Folder data. Aborting...");
			e.printStackTrace();
			return;
		}
		for(MailFolder f: folders.getCurrentPage()) {
			//logger.info(f.displayName + "\t" + f.id);
			if(f.displayName.equals("Processed Warrants")) {
				mailboxFolderId = f.id;
				logger.info("Folder Name: " + f.displayName + "\t" + mailboxFolderId);
			}
		}
		
		//Getting the Emails!
		MessageCollectionPage searchMailboxEmails;
		try {
			//TODO: Uncomment below line, and comment the line below that
			searchMailboxEmails = Graph.listMailFromDistributionList(mailboxAddress);
			//searchMailboxEmails = Graph.listMailFromDistributionList(mailboxAddress, "AAMkADI3MzkwNDNmLTliODItNDFjOC1iYjJiLTc0MTM5MzQ3NDQxYwAuAAAAAAD0Ma_rXLN5RoTcEAl_Shz8AQDXylSq6leYT4rYCHzhVtl6AAAIzwbUAAA=");
		} catch (Exception e) {
			logger.error("Error in searching mailbox! Aborting...");
			e.printStackTrace();
			return;
		}
		if(searchMailboxEmails.getCurrentPage().size() == 0) {
			logger.info("No new warrant emails! Ending execution...");
			return;
		}
		
		for(Message m: searchMailboxEmails.getCurrentPage()) {
			sqlRow = new ArrayList<>();
			ASAEmail = "";
			warrantControlNumber = "";
			sequenceNumber = "";
			yesNo = "";
			fwd = false;
			noWarrantFlag = false;
			
			mailSubject = m.subject;
			
			logger.info("Current email subject: \t" + mailSubject);
			fwdMatcher = fwdPattern.matcher(mailSubject);
			while(fwdMatcher.find()) {
				logger.info("Detected forwarded email! Editing effective subject...");
				mailSubject = mailSubject.replace(fwdMatcher.group(), "").trim();
				fwd = true;
			}
			
			if(fwd == true) {
				logger.info("Edited effective email subject: \t" + mailSubject);
			}
			
			subjectTokens = mailSubject.split("[\\s ]+");
			/* 
			for(String temp: subjectTokens) {
				//System.out.println(temp);
			} 
			*/
			warrantControlNumber = subjectTokens[0];
			sequenceNumber = subjectTokens[1];
			ASAEmail = m.sender.emailAddress.address;
			
			if(subjectTokens[2].toLowerCase().equals("yes")) {
				yesNo = "Y";
			} else {
				yesNo = "N";
			}
			
			if(warrantControlNumber.toLowerCase().equals("nowarrant")) {
				logger.info("No Warrant email detected!");
				noWarrantFlag = true;
			}
			
			//logger.info(warrantControlNumber + "\t" + sequenceNumber + "\t" + ASAEmail + "\t" + yesNo);
			
			AttachmentCollectionPage attachments = null;
			try {
				attachments = Graph.getMailAttachments(mailboxAddress, m.id);
			} catch (Exception e) {
				logger.error("Error fetching attachments! Moving to next email...");
				e.printStackTrace();
				continue;
			}
			
			//Setting the attachment counter to zero for each Email
			validAttachments = 0;
			
			for(Attachment attachment: attachments.getCurrentPage()) {
				fileAttachment = null;
				if(attachment.contentType == null) {
					logger.info("Null Content Type. Moving to next attachment...");
					continue;
				}
				
				fileAttachment = (FileAttachment) attachment;
				
				if(fileAttachment.contentBytes.length == 0) {
					logger.info(fileAttachment.name + " has content size of " + fileAttachment.contentBytes.length + " bytes. Moving to next attachment...");
					continue;
				}
				

				attachmentTokens = attachment .name.split("\\.");
				len = attachmentTokens.length;
				
				if(attachmentTokens[len-1].equals("png")) {
					logger.info("Attachment is an image. Ignoring ...");
					continue;
				}
				
				logger.info("Downloading email attachment: " + attachment.name);
				try {
					Graph.downloadAttachment_test2(mailboxAddress, m.id, attachment, downloadDirectory, warrantControlNumber + " - " + attachment.name);
				} catch (Exception e) {
					logger.error("Error downloading attachment. Moving to next attachment...");
					e.printStackTrace();
					continue;
				}
				
				validAttachments += 1;
				
			}
			
			if((validAttachments >= 1) || (noWarrantFlag == true && validAttachments == 0)) {
				sqlRow.add(warrantControlNumber);
				sqlRow.add(ASAEmail);
				sqlRow.add(sequenceNumber);
				sqlRow.add(yesNo);
				
				insertedRows.add(sqlRow);
				
				try {
					Graph.moveEmail(mailboxAddress, m.id, mailboxFolderId);
				} catch (Exception e) {
					logger.error("Error moving email into \"Processed Warrants\" folder. Moving to next email...");
					e.printStackTrace();
					continue;
				}
				
			} else {
				logger.info("The mail has no valid warrant documents. Continuing...");
				continue;
			}
			
		}
		
		if(insertedRows.size() == 0) {
			logger.info("No new rows to update DB. Ending execution...");
			return;
		}
		
		String insertQuery = getInsertQuery(insertedRows);
		//logger.info(insertQuery);
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			String connectionURL = "jdbc:sqlserver://SAOJD1;"
					+ "databaseName=PD51_Data;"
					+ "integratedSecurity=true";
			Connection conn = DriverManager.getConnection(connectionURL);
			PreparedStatement ps;
			
			ps = conn.prepareStatement(insertQuery);
			ps.execute();
			logger.info("updated _bcsao_WarrantControl table with " + insertedRows.size() + " rows.");
			
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		
		
	}

}

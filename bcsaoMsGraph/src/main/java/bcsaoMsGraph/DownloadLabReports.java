package bcsaoMsGraph;

import java.io.File;
import java.io.IOException;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.microsoft.graph.models.Attachment;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.requests.AttachmentCollectionPage;
import com.microsoft.graph.requests.MessageCollectionPage;

/*
 * https://github.com/rutvikdixit/bcsaoMsGraph/
 */
public class DownloadLabReports {
	public static void initializeGraph(Properties properties) {
		try {
			Graph.initalizeGraphforApp(properties);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	public static void main(String[] args) {
		// TODO Auto-generated method stub

		Logger logger = LoggerFactory.getLogger("BPDDownloadsLogger");

		Properties oAuth = new Properties();
		try {
			oAuth.load(DownloadWarrants.class.getResourceAsStream("/oAuth.properties"));
		} catch (IOException e) {
			e.printStackTrace();
		}
		initializeGraph(oAuth);
		//logger.info(oAuth.toString());
		
		File downloadDirectory = new File("C:\\DownloadsBPD");
		
		if(!downloadDirectory.exists()) {
			logger.info("Download Directory not found. Creating now ...");
			downloadDirectory.mkdir();
			logger.info("Download Directory created.");
		} else {
			logger.info("Download Directory found.");
		}
		

		Pattern p = Pattern.compile("\\d{9}");
		Matcher matcher = null;
		
		
		try {
			
			/*
			 * "Processed" Outlook FolderID AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQAuAAAAAACdE_6qmw7pT5dfvbw3JIQnAQCmFx3FC2puSorHsETnN4atAAAK-ownAAA=
			 */
			
			/*
			MailFolderCollectionPage page =  Graph.getFolders(mailboxAddress);
			for(MailFolder folder: page.getCurrentPage()) {
				logger.info(folder.displayName + "\t\t\t" + folder.id);
			}
			*/
			String mailboxAddress = "saobpdlabs@stattorney.org";
			String bpdSenderAddress = "fsl-no-reply@baltimorepolice.org";
			String moveFolderID = "AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQAuAAAAAACdE_6qmw7pT5dfvbw3JIQnAQCmFx3FC2puSorHsETnN4atAAAK-ownAAA=";
			String subject = "", ccNumber = "", newFileName;
			String[] attachmentTokens;
			int len;
			
			
			//Retrieving all messages from the inbox
			MessageCollectionPage messages = Graph.searchInbox(mailboxAddress, bpdSenderAddress);
			
			if(messages.getCurrentPage().size() == 0) {
				logger.info("No new lab reports received!");
			}
			
			for(Message m: messages.getCurrentPage()) {
				//Resetting variables for iteration
				subject = "";
				ccNumber = "";
				newFileName = "";
				//Ignore emails which are not from fsl-no-reply@baltimorepolice.org
				
				//Assigning 
				subject = m.subject;
				
				matcher = p.matcher(subject);
				while(matcher.find()) {
					ccNumber = matcher.group();
				}
				logger.info(ccNumber);
				
				//Downloading the attachments
				AttachmentCollectionPage attachments =  Graph.getMailAttachments(mailboxAddress, m.id);
				for(Attachment a: attachments.getCurrentPage()) {
					//Skipping the attachment if the extension is not .pdf
					len = 0;
					attachmentTokens = a.name.split("\\.");
					len = attachmentTokens.length;
					if(!attachmentTokens[len - 1].toLowerCase().equals("pdf")) {
						logger.info("Attachment is not a PDF: \t\t" + a.name);
						continue;
					}
					
					
					
					newFileName = ccNumber + " - " + a.name;
					logger.info(newFileName);
					Graph.downloadAttachment_test2(mailboxAddress, m.id, a, downloadDirectory, newFileName);
					
				}
				
				//Code to move email to the Processed folder
				Graph.moveEmail(mailboxAddress, m.id, moveFolderID);
				
				
				
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
	}

}

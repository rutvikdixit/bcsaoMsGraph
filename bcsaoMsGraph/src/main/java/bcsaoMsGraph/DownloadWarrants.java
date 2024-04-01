package bcsaoMsGraph;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.microsoft.graph.models.Attachment;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.requests.AttachmentCollectionPage;
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
		
		String fileName = "";
		File downloadDirectory = new File("Download");
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
			
			MessageCollectionPage searchMailBox;
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}

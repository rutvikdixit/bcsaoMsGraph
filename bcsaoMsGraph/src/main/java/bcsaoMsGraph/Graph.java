package bcsaoMsGraph;


import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.azure.core.credential.AccessToken;
import com.azure.core.credential.TokenRequestContext;
import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.Attachment;
import com.microsoft.graph.models.MailFolder;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.models.MessageMoveParameterSet;
import com.microsoft.graph.requests.AttachmentCollectionPage;
import com.microsoft.graph.requests.FileAttachmentRequestBuilder;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.MailFolderCollectionPage;
import com.microsoft.graph.requests.MessageCollectionPage;
import okhttp3.Request;

public class Graph {

    private static Properties _properties;
    private static ClientSecretCredential _clientSecretCredential;
    private static GraphServiceClient<Request> _appClient;
    private static Logger logger;
    
    public static void initalizeGraphforApp(Properties properties) throws Exception {
    	logger = LoggerFactory.getLogger("GraphLogger");
    	
    	if(properties == null) {
    		throw new Exception("Properties cannot be null");
    	}
    	
    	_properties = properties;
    	
    	if(_clientSecretCredential == null) {
    		final String clientId, clientSecret, tenantId;
    		
    		clientId = _properties.getProperty("app.clientId");
    		clientSecret = _properties.getProperty("app.clientSecret");
    		tenantId = _properties.getProperty("app.tenantId");
    		
    		_clientSecretCredential = new ClientSecretCredentialBuilder()
    				.clientId(clientId)
    				.clientSecret(clientSecret)
    				.tenantId(tenantId)
    				.build();
    	}
    	
    	if(_appClient == null) {
    		final TokenCredentialAuthProvider authProvider = new TokenCredentialAuthProvider(List.of("https://graph.microsoft.com/.default"), _clientSecretCredential) ;
    		
    		_appClient = GraphServiceClient.builder()
    				.authenticationProvider(authProvider)
    				.buildClient();
    	}
    	
    }
    
    public static MessageCollectionPage listMail() throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }

    	MessageCollectionPage page2 = _appClient.users("saobpdlabs@stattorney.org").mailFolders("inbox").messages().buildRequest().top(10).get();
    	//MessageCollectionPage page2 = _appClient.users("saobpdlabs@stattorney.org").messages().buildRequest().select("subject, from").top(15).get();
    	return page2;
    	
    }
    
    public static MessageCollectionPage listMailFromDistributionList(String emailAddress) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        } 
    	
    	if(emailAddress == null || emailAddress == "") {
    		throw new Exception("Email Address cannot be null or empty string.");
    	}
    	

    	MessageCollectionPage page2 = _appClient.users(emailAddress).mailFolders("inbox").messages().buildRequest().top(10).get();
    	
    	return page2;
    }
    
    public static MessageCollectionPage listMailFromDistributionList(String emailAddress, String folderID) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        } 
    	
    	if(emailAddress == null || emailAddress == "") {
    		throw new Exception("Email Address cannot be null or empty string.");
    	}
    	

    	MessageCollectionPage page2 = _appClient.users(emailAddress)
    			.mailFolders(folderID)
    			.messages()
    			.buildRequest()
    			.top(10)
    			.get();
    	
    	return page2;
    }
        
    
    public static MessageCollectionPage listMail(String emailAddress) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        } 
    	
    	if(emailAddress == null || emailAddress == "") {
    		throw new Exception("Email Address cannot be null or empty string.");
    	}

    	MessageCollectionPage page2 = _appClient.users(emailAddress).messages(). buildRequest(). top(10).get();
    	//MessageCollectionPage page2 = _appClient.users("saobpdlabs@stattorney.org").messages().buildRequest().select("subject, from").top(15).get();
    	return page2;
    	
    }
    
    public static MailFolderCollectionPage getFolders(String mailboxEmail) throws Exception {

    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	MailFolderCollectionPage folders = _appClient.users(mailboxEmail)
    	.mailFolders()
    	.buildRequest()
    	.top(20)
    	.get();
    	
    	return folders;
    	
    }
    
    public static void moveEmail(String mailBox, String emailId, String folderId) throws Exception {

    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	if((mailBox == null) || (emailId == null) || (folderId == null)) {
    		throw new Exception("MailBox/EmailID/FolderID cannot be null");
    	}
    	
    	_appClient.users(mailBox)
    	.messages(emailId)
    	.move(MessageMoveParameterSet.newBuilder().withDestinationId(folderId).build())
    	.buildRequest()
    	.post();
    	
    	
    }
    
    public static MessageCollectionPage searchInbox(String searchMailbox, String senderEmail) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	//MessageCollectionPage searchResult = _appClient.users(searchMailbox).mailFolders("inbox").messages().buildRequest().filter("sender/emailAddress/address eq '" + senderEmail + "' and contains(subject, 'graph')").get();
    	MessageCollectionPage searchResult = _appClient.users(searchMailbox)
    			.mailFolders("inbox")
    			.messages()
    			.buildRequest()
    			.filter("receivedDateTime ge 2022-01-01T00:00:00Z and sender/emailAddress/address eq '" + senderEmail + "'")
    			.orderBy("receivedDateTime DESC")
    			.top(20)
    			.get();
    	
    	return searchResult;
    }
    

    public static AttachmentCollectionPage getMailAttachments(String searchMailbox, String messageId) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	AttachmentCollectionPage attachments = _appClient.users(searchMailbox).messages(messageId).attachments().buildRequest().get();
    	
		return attachments;
    }
    
    public static File downloadAttachment_test2(String mailBox, String mailId, Attachment attachment, File downloadDirectory, String optionalFileName) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	if((mailBox == null) || (mailId == null) || (attachment == null) || (downloadDirectory == null)) {
    		throw new Exception("MailBox/MailID/Attachment/DownloadDirectory cannot be null");
    	}
    	
    	if(!downloadDirectory.isDirectory()) {
    		throw new Exception("Supplied DownloadDirectory parameter is not a directory");
    	}
    	
    	String req = _appClient.users(mailBox)
    			.messages(mailId)
    			.attachments(attachment.id)
    			.getRequestUrl();
    	
    	logger.info(req);
    	
    	String fileName;
    	if(optionalFileName == null || optionalFileName == "") {
    		fileName = attachment.name;
    		
    	} else {
    		fileName = optionalFileName;
    	}
    	
    	FileAttachmentRequestBuilder farb = new FileAttachmentRequestBuilder(req, _appClient, null);
    	InputStream is = farb.content().buildRequest().get();
    	
    	File downloadFile = new File(fileName);
    	File finalFile = new File(downloadDirectory.getPath() + "\\" + fileName);
    	
    	Files.copy(is, downloadFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
    	Files.move(downloadFile.toPath(), finalFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
    	
    	return null;
    }
    
    public static File downloadAttachment_test1(String mailBox, String mailId, Attachment attachment, String optionalFileName) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	if((mailBox == null) || (mailId == null) || (attachment == null)) {
    		throw new Exception("MailBox/MailID/Attachment cannot be null");
    	}
    	
    	String req = _appClient.users(mailBox)
    			.messages(mailId)
    			.attachments(attachment.id)
    			.getRequestUrl();
    	
    	logger.info(req);
    	
    	String fileName;
    	if(optionalFileName == null || optionalFileName == "") {
    		fileName = attachment.name;
    		
    	} else {
    		fileName = optionalFileName;
    	}
    	
    	FileAttachmentRequestBuilder farb = new FileAttachmentRequestBuilder(req, _appClient, null);
    	InputStream is = farb.content().buildRequest().get();
    	
    	File tempDir = Files.createTempDirectory("MSGraph").toFile();
    	File downloadFile = new File(fileName);
    	File finalFile = new File(tempDir.getPath() + "\\" + fileName);
    	
    	Files.copy(is, downloadFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
    	Files.move(downloadFile.toPath(), finalFile.toPath(), StandardCopyOption.ATOMIC_MOVE);
    	
    	logger.info(tempDir.getPath());
    	
    	
    	return null;
    }
    
    public static AttachmentCollectionPage _dontUseThis_getAttachments(String searchMailbox, String messageId) throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	String mId = null;
    	
    	MessageCollectionPage searchResult = _appClient.users(searchMailbox).mailFolders("inbox").messages().buildRequest().filter("sender/emailAddress/address eq 'rdixit@stattorney.org' and contains(subject, 'graph')").get();
    	
    	for(Message m: searchResult.getCurrentPage()) {
    		logger.info(m.subject + "\t" + m.id);
    		mId = m.id;
    	}
    	
    	AttachmentCollectionPage attachResult = _appClient.users(searchMailbox).messages(mId).attachments().buildRequest().get();
    	
    	for(Attachment a: attachResult.getCurrentPage()) {
    		logger.info(a.name + "\t" + a.contentType + "\t" + a.id);
    	}
    	
    	
    	String req = _appClient.users(searchMailbox).messages(mId).attachments("AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD9y4DqAAABEgAQAIge9hDx29lGsF8gThfzhgI=").getRequestUrl();
    	logger.info(req);
    	
    	FileAttachmentRequestBuilder farb = new FileAttachmentRequestBuilder(req, _appClient, null);
    	
    	InputStream is = farb.content().buildRequest().get();
    	File f = new File("down.pdf");
    	Files.copy(is, f.toPath(), StandardCopyOption.REPLACE_EXISTING);
    	
    	is.close();
    	
    	
    	/*
    	
    	// SAMPLE attachmentId AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD9y4DqAAABEgAQAIge9hDx29lGsF8gThfzhgI=
    	FileAttachment fileAttach = (FileAttachment) _appClient.users(searchMailbox).messages(mId).attachments("AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD9y4DqAAABEgAQAIge9hDx29lGsF8gThfzhgI=").buildRequest().get();
    	String req = _appClient.users(searchMailbox).messages(mId).attachments("AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD9y4DqAAABEgAQAIge9hDx29lGsF8gThfzhgI=").getRequestUrl();
    	logger.info(req);
    	logger.info(fileAttach.name);
    	logger.info(fileAttach.size.toString());
    	
    	byte[] fileEncodedBytes = fileAttach.contentBytes;
    	System.out.println(fileEncodedBytes.toString());
    	
    	byte[] fileDecodedBytes = Base64.getMimeDecoder().decode(fileEncodedBytes);
    	
    	File download = new File(fileAttach.name); 
    	System.out.println(fileDecodedBytes.length);
    	OutputStream os = new FileOutputStream(download);
    	os.write(fileDecodedBytes);
    	os.close();
    	*/
    	
    	
    	return null;
    }
    
    public static void _dontUseThis_moveEmail() throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }
    	
    	Message message = _appClient.users("saobpdlabs@stattorney.org")
    			.messages("AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD5EonTAAA=")
    			.buildRequest()
    			.get();
    	
    	System.out.println(message.changeKey);
    	
    	
    	MailFolderCollectionPage folders = _appClient.users("saobpdlabs@stattorney.org")
    	.mailFolders()
    	.buildRequest()
    	.get();
    	
    	for(MailFolder f: folders.getCurrentPage()) {
    		System.out.println(f.displayName + "\t" + f.id);
    	}
    	
    	//"MSGraph" folderId: AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQAuAAAAAACdE_6qmw7pT5dfvbw3JIQnAQCmFx3FC2puSorHsETnN4atAAD5Ep_zAAA=
    	//TestAPI EmailId: AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD5EonTAAA=
    	
    	Message message2 = _appClient.users("saobpdlabs@stattorney.org")
		.messages("AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQBGAAAAAACdE_6qmw7pT5dfvbw3JIQnBwCmFx3FC2puSorHsETnN4atAAAAAAEMAACmFx3FC2puSorHsETnN4atAAD5EonTAAA=")
    	.move(MessageMoveParameterSet.newBuilder().withDestinationId("AAMkAGI4MjljYzc4LTQzNDAtNDFiOS05ODAwLWQ5ODI3MzA1N2Y5OQAuAAAAAACdE_6qmw7pT5dfvbw3JIQnAQCmFx3FC2puSorHsETnN4atAAD5Ep_zAAA=").build())
    	.buildRequest()
    	.post();
    	
    	System.out.println(message2.changeKey);
    	
    }
    
    public static String getToken() throws Exception {
    	if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }

        // Request the .default scope as required by app-only auth
        final String[] graphScopes = new String[] {"https://graph.microsoft.com/.default"};

        final TokenRequestContext context = new TokenRequestContext();
        context.addScopes(graphScopes);

        final AccessToken token = _clientSecretCredential.getToken(context).block();
        return token.getToken();
    }

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		System.out.println("Hello World!");
		
	}

}


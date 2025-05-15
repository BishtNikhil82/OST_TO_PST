package com.NAtools.Convertor;

import com.NAtools.config.LogManagerConfig;
import com.NAtools.util.OfficeConnectionUtility;
import com.aspose.email.*;
import java.util.*;
import java.util.logging.Logger;

public class OSTToOffice365 {
    private static final Logger logger = Logger.getLogger(OSTToOffice365.class.getName());

    static {
        LogManagerConfig.configureLogger(logger);
    }

    private final PersonalStorage sourcePst;
    public final IEWSClient exchangeClient;
    private final boolean enableDuplicateCheck;
    private final List<String> listduplicacy = new ArrayList<>();
    private final List<String> listdupliccal = new ArrayList<>();
    private final List<String> listduplictask = new ArrayList<>();
    private final List<String> listdupliccontact = new ArrayList<>();
    public final Map<String, FolderInfo> loadedFolders = new HashMap<>();

    public OSTToOffice365(PersonalStorage sourcePst, String username, String password, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.enableDuplicateCheck = enableDuplicateCheck;
        this.exchangeClient = connectToOffice365(username, password);
    }

    private IEWSClient connectToOffice365(String username, String password) {
        IEWSClient client = null;
        try {
            // Attempt to connect using OAuth2 first
            client = OfficeConnectionUtility.connectToOffice365();
        } catch (Exception e) {
            logger.severe("OAuth2 connection failed: " + e.getMessage() + ". Trying basic authentication.");
            // If OAuth2 connection fails, fallback to basic authentication
            client = connectToOffice365Basic(username, password);
        }
        return client;
    }

    private IEWSClient connectToOffice365Basic(String username, String password) {
        IEWSClient client = null;
        while (client == null) {
            try {
                client = EWSClient.getEWSClient("https://outlook.office365.com/EWS/Exchange.asmx", username, password);
                EmailClient.setSocketsLayerVersion2(true);
                client.setTimeout(300000);
                EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
                logger.info("Connected to Office 365 using basic authentication.");
            } catch (Exception e) {
                logger.severe("Failed to connect to Office 365 using basic authentication. Retrying...");
                e.printStackTrace();
            }
        }
        return client;
    }

    // Load initial folders from the OST file and replicate them in Office 365
    public void loadInitialFolders() {
        FolderInfo rootFolder = sourcePst.getRootFolder();
        processFolder(rootFolder, exchangeClient.getMailboxInfo().getRootUri());
        logger.info("Initial folders loaded and created in Office 365.");
    }

    // Process each folder from the OST file and replicate its structure in Office 365
    private void processFolder(FolderInfo sourceFolder, String parentFolderUri) {
        try {
            String folderName = cleanFolderName(sourceFolder.getDisplayName());
            if (folderName.isEmpty()) {
                logger.warning("Empty folder name. Skipping folder creation.");
                return;
            }

            String folderUri = exchangeClient.createFolder(parentFolderUri, folderName).getUri();
            logger.info("Created folder in Office 365: " + folderName);

            processMessages(sourceFolder, folderUri);

            for (FolderInfo subFolder : sourceFolder.getSubFolders()) {
                processFolder(subFolder, folderUri);
            }
        } catch (Exception e) {
            logger.severe("Error processing folder: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // Process messages and send them to the corresponding folder in Office 365
    private void processMessages(FolderInfo sourceFolder, String folderUri) {
        try {
            MessageInfoCollection messages = sourceFolder.getContents();
            for (int i = 0; i < messages.size(); i++) {
                try {
                    MessageInfo messageInfo = messages.get_Item(i);
                    MapiMessage message = sourcePst.extractMessage(messageInfo);
                    processMessage(message, folderUri);
                    logger.info("Processed message: " + message.getSubject());
                } catch (Exception e) {
                    logger.warning("Error processing message: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing messages: " + e.getMessage());
        }
    }

    // Process individual messages and send them to Office 365
    private void processMessage(MapiMessage message, String folderUri) {
        try {
            switch (message.getMessageClass()) {
                case "IPM.Appointment":
                case "IPM.Schedule.Meeting.Request":
                    processCalendar(message, folderUri);
                    break;
                case "IPM.Contact":
                    processContact(message, folderUri);
                    break;
                case "IPM.Task":
                    processTask(message, folderUri);
                    break;
                default:
                    processGenericMessage(message, folderUri);
                    break;
            }
        } catch (Exception e) {
            logger.warning("Error processing message: " + e.getMessage());
        }
    }

    private void processCalendar(MapiMessage message, String folderUri) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            if (!isDuplicateCalendar(calendar)) {
                exchangeClient.appendMessage(folderUri, message.toMailMessage(new MailConversionOptions()));
            }
        } catch (Exception e) {
            logger.warning("Error processing calendar: " + e.getMessage());
        }
    }

    private void processContact(MapiMessage message, String folderUri) {
        try {
            MapiContact contact = (MapiContact) message.toMapiMessageItem();
            if (!isDuplicateContact(contact)) {
                exchangeClient.createContact(folderUri, Contact.to_Contact(contact));
            }
        } catch (Exception e) {
            logger.warning("Error processing contact: " + e.getMessage());
        }
    }

    private void processTask(MapiMessage message, String folderUri) {
        try {
            MapiTask task = (MapiTask) message.toMapiMessageItem();
            if (!isDuplicateTask(task)) {
                ExchangeTask exchangeTask = new ExchangeTask();
                exchangeTask.setSubject(task.getSubject());
                exchangeTask.setBody(task.getBody());
                exchangeTask.setStartDate(task.getStartDate());
                exchangeTask.setDueDate(task.getDueDate());
                exchangeTask.setStatus(task.getStatus());
// Copy other necessary properties from MapiTask to ExchangeTask
                exchangeClient.createTask(folderUri, exchangeTask);
            }
        } catch (Exception e) {
            logger.warning("Error processing task: " + e.getMessage());
        }
    }

    private void processGenericMessage(MapiMessage message, String folderUri) {
        try {
            if (!isDuplicateMessage(message)) {
                exchangeClient.appendMessage(folderUri, message.toMailMessage(new MailConversionOptions()));
            }
        } catch (Exception e) {
            logger.severe("Error processing generic message: " + e.getMessage());
        }
    }

    // Duplicate check methods
    private boolean isDuplicateCalendar(MapiCalendar calendar) {
        String key = calendar.getLocation() + calendar.getStartDate() + calendar.getEndDate();
        if (listdupliccal.contains(key)) {
            return true;
        } else {
            listdupliccal.add(key);
            return false;
        }
    }

    private boolean isDuplicateContact(MapiContact contact) {
        String key = contact.getNameInfo().getDisplayName() + contact.getPersonalInfo().getNotes();
        if (listdupliccontact.contains(key)) {
            return true;
        } else {
            listdupliccontact.add(key);
            return false;
        }
    }

    private boolean isDuplicateTask(MapiTask task) {
        String key = task.getSubject() + task.getBody();
        if (listduplictask.contains(key)) {
            return true;
        } else {
            listduplictask.add(key);
            return false;
        }
    }

    private boolean isDuplicateMessage(MapiMessage message) {
        String key = message.getSubject() + message.getBody();
        if (listduplicacy.contains(key)) {
            return true;
        } else {
            listduplicacy.add(key);
            return false;
        }
    }

    private String cleanFolderName(String folderName) {
        if (folderName == null || folderName.trim().isEmpty()) {
            return "";
        }
        return folderName.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    public static void main(String[] args) {
        String ostFilePath = "your_ost_file_path.ost";
        String username = "your_username@domain.com";
        String password = "your_password";

        PersonalStorage sourcePst = PersonalStorage.fromFile(ostFilePath);
        OSTToOffice365 converter = new OSTToOffice365(sourcePst, username, password, true);
        converter.loadInitialFolders();
    }
}

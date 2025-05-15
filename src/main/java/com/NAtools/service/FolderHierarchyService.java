package com.NAtools.service;

import com.NAtools.config.*;
import com.NAtools.db.DBConnection;
import com.NAtools.db.DBOperations;
import com.NAtools.db.DatabaseHelper;
import com.aspose.email.*;
import com.aspose.email.system.collections.generic.IGenericEnumerator;
import com.aspose.email.system.collections.generic.KeyValuePair;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Calendar;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.logging.Logger;

public class FolderHierarchyService implements Runnable {
    private static final Logger logger = Logger.getLogger(FolderHierarchyService.class.getName());
    static {
        LogManagerConfig.configureLogger(logger);
        logger.info("Test log from AnotherClass.");
    }
    private final PersonalStorage sourcePst;
    private final PersonalStorage targetPst;
    private final String sourcePstPath;
    private final List<String> listduplicacy = new ArrayList<>();
    private final List<String> listdupliccal = new ArrayList<>();
    private final List<String> listduplictask = new ArrayList<>();
    private final List<String> listdupliccontact = new ArrayList<>();
    private long count_destination = 0;
    private final boolean enableDuplicateCheck;
    private List<FolderInfo> loadedFolders = new ArrayList<>();

    private final CountDownLatch latch;
    private FolderInfo lostAndFoundFolder;
    private DBOperations dbOps;

    public FolderHierarchyService(PersonalStorage sourcePst, PersonalStorage targetPst, String sourcePstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.targetPst = targetPst;
        this.sourcePstPath = sourcePstPath;
        this.enableDuplicateCheck = enableDuplicateCheck;
        this.latch = new CountDownLatch(100);
        dbOps = new DBOperations();
        DatabaseHelper.executeSqlScript();
    }

    @Override

    public void run() {
        Connection conn = null;
        try {
            logger.info("Attempting to establish database connection...");
            conn = DBConnection.getInstance().getConnection(); // Get a connection instance for this thread

            if (conn != null && !conn.isClosed()) {
                logger.info("FolderHierarchyService Database connection established successfully.");
                processFolders(conn);
            } else {
                logger.severe("Failed to establish database connection.");
            }

            logger.info("Target PST file saved successfully.");
        } catch (SQLException e) {
            logger.severe("Error during database operations: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try {
                if (conn != null && !conn.isClosed()) {
                    conn.close(); // Close the connection specific to this thread
                    logger.info("Database connection closed.");
                }
            } catch (SQLException ex) {
                logger.severe("Error closing database connection: " + ex.getMessage());
                ex.printStackTrace();
            }
            sourcePst.dispose(); // Ensure PST files are properly closed
            targetPst.dispose();
        }
    }




    public void processFolders(Connection conn) {
        try {
            FolderInfo sourceRootFolder = sourcePst.getRootFolder();
            FolderInfo targetRootFolder = targetPst.getRootFolder();
            lostAndFoundFolder = createLostAndFoundFolder(targetRootFolder);
            logger.info("Root Folder: " + sourcePstPath);

            // Process the root folder
            processFolder(conn, sourceRootFolder, targetRootFolder, sourcePstPath, null); // Pass null as parentId for root

        } catch (Exception e) {
            logger.severe("An error occurred while processing folders: " + e.getMessage());
            e.printStackTrace();
        }
    }


    private FolderInfo createLostAndFoundFolder(FolderInfo rootFolder) {
        for (FolderInfo folder : rootFolder.enumerateFolders()) {
            if (folder.getDisplayName().equalsIgnoreCase("Lost and Found")) {
                return folder;
            }
        }
        return rootFolder.addSubFolder("Lost and Found");
    }

    private void processFolder(Connection conn, FolderInfo sourceFolder, FolderInfo targetFolder, String parentPath, Integer parentId) {
        try {
            String folderName = cleanFolderName(sourceFolder.getDisplayName());
            if (folderName.isEmpty()) {
                folderName = generateFolderNameFromPath(parentPath);
                logger.info("Unnamed folder found. Generated name: " + folderName);
            }

            String newFolderName = folderName;
            if (parentPath.endsWith("IPM_SUBTREE")) {
                // Calculate the message count if the parent is IPM_SUBTREE
                int messageCount = sourceFolder.getContents().size();
                newFolderName = folderName + " [" + messageCount + "]";
            }


            String fullPath = parentPath.isEmpty() ? newFolderName : parentPath + "/" + newFolderName;
            //logger.info("Processing folder: " + newFolderName + " fullPath: " + fullPath + " contains " + sourceFolder.getContents().size() + " messages.");

            // Create the new folder in the target PST
            FolderInfo newTargetFolder = targetFolder.addSubFolder(newFolderName, true);

            // Insert folder into the database

            int newFolderId = dbOps.insertFolder(conn,folderName, parentId);

            // Process subfolders recursively
            FolderInfoCollection subFolders = sourceFolder.getSubFolders();
            for (FolderInfo subFolder : subFolders) {
                processFolder(conn,subFolder, newTargetFolder, fullPath, newFolderId);
            }

            // Process messages in the current folder
            processMessages(conn,sourceFolder, newTargetFolder, newFolderId);

        } catch (Exception e) {
            logger.severe("Error processing folder: " + e.getMessage());
        }
    }


    private boolean isOrphaned(MapiMessage message) {

        if (!"IPM.Note".equals(message.getMessageClass())) {
            // Skip processing for non-IPM.Note messages
            return false;
        }
        MapiProperty parentFolderIdProp = message.getProperties().get_Item(MapiPropertyTag.PR_PARENT_ENTRYID);
        if (parentFolderIdProp == null || parentFolderIdProp.getData() == null) {
            // Log all properties of the message
            IGenericEnumerator<KeyValuePair<Long, MapiProperty>> iterator = message.getProperties().iterator();
            while (iterator.hasNext()) {
                KeyValuePair<Long, MapiProperty> entry = iterator.next();
                Long propertyKey = entry.getKey();
                MapiProperty property = entry.getValue();
                byte[] propertyData = property.getData();

                if (propertyData != null) {
                    // Convert the byte array to a hexadecimal string
                    StringBuilder hexString = new StringBuilder();
                    for (byte b : propertyData) {
                       hexString.append(String.format("%02X", b));
                    }

                    logger.info("Property: " + propertyKey + ", Value: " + hexString.toString());
                } else {
                    logger.info("Property: " + propertyKey + ", Value: null");
                }
            }

            logger.warning("Orphaned email detected: Missing or invalid PR_PARENT_ENTRYID. " +
                    "Subject: " + message.getSubject() +
                    ", Sender: " + message.getSenderEmailAddress() +
                    ", Message Class: " + message.getMessageClass() +
                    ", Message ID: " + message.getItemId());

            return true;
        }
        logger.warning("Should not reach there");
        // Log the PR_PARENT_ENTRYID property data
//        byte[] parentFolderIdData = parentFolderIdProp.getData();
//        if (parentFolderIdData != null) {
//            StringBuilder hexString = new StringBuilder();
//            for (byte b : parentFolderIdData) {
//                hexString.append(String.format("%02X", b));
//            }
//
//            logger.info("PR_PARENT_ENTRYID: " + hexString.toString() +
//                    " for message with subject: " + message.getSubject() +
//                    ", Sender: " + message.getSenderEmailAddress());
//        } else {
//            logger.warning("PR_PARENT_ENTRYID data is null for message with subject: " + message.getSubject());
//        }

        return false;
    }

    private void processMessages(Connection conn, FolderInfo sourceFolder, FolderInfo targetFolder, int folderId) {
        try {
            MessageInfoCollection messages = sourceFolder.getContents();

            for (int i = 0; i < messages.size(); i++) {
                try {
                    MessageInfo messageInfo = messages.get_Item(i);
                    MapiMessage message = sourcePst.extractMessage(messageInfo);

                    Date deliveryTime = null;

                    // Check if the message is a draft
                    if (message.getMessageClass().equals("IPM.Note")) {
                        deliveryTime = message.getDeliveryTime();
                    }

                    // If no delivery time, check other potential date fields
                    if (deliveryTime == null) {
                        MapiProperty submitTimeProp = message.getProperties().get_Item(MapiPropertyTag.PR_CLIENT_SUBMIT_TIME);
                        if (submitTimeProp != null) {
                            deliveryTime = submitTimeProp.getDateTime();
                        }
                    }

                    // If no date was found, set a fallback date (e.g., current date)
                    if (deliveryTime == null) {
                        deliveryTime = new Date();
                        logger.warning("No date found for draft; using current date.");
                    }

                    // Convert the date to the required format
                    SimpleDateFormat isoFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    String receivedDateString = isoFormat.format(deliveryTime);
                    Timestamp receivedDate = Timestamp.valueOf(receivedDateString);

                    logger.info("Converted date ISO format = " + receivedDateString);

                    // Get the sender's SMTP email address
                    String senderEmail = message.getSenderEmailAddress();
                    MapiProperty senderSmtpAddressProp = message.getProperties().get_Item(MapiPropertyTag.PR_SENDER_SMTP_ADDRESS_W);

                    if (senderSmtpAddressProp != null) {
                        senderEmail = senderSmtpAddressProp.getString();
                    }
                    if (senderEmail == null || senderEmail.trim().isEmpty()) {
                        senderEmail = "No Sender"; // Fallback
                        logger.warning("Sender email is null or empty. Using fallback.");
                    }

                    // Get the recipients' email addresses (To)
                    String recipients = getRecipientEmails(message);
                    String bodyFormat;
                    String bodyContent;

                    if (message.getBodyHtml() != null) {
                        bodyFormat = "HTML";
                        bodyContent = message.getBodyHtml();
                    } else if (message.getBodyRtf() != null) {
                        bodyFormat = "RTF";
                        bodyContent = message.getBodyRtf();
                    } else {
                        bodyFormat = "PlainText";
                        bodyContent = message.getBody();
                    }
                    // Insert the message into the database and get the generated message ID
                    String subject = message.getSubject() != null ? message.getSubject() : "No Subject"; // Fallback for null subjects
                    int messageId = dbOps.insertMessage(conn, folderId, subject, senderEmail, bodyContent, receivedDateString, recipients,bodyFormat);

                    // Process and save attachments to the database
                    if (messageId > 0) {
                        if (!message.getAttachments().isEmpty()) {
                            saveAttachmentsToDatabase(conn, message, messageId);
                        }
                    }

                    // Process the message (e.g., add to the target PST)
                    processMessage(targetFolder, message);

                    // Log preview information
                    logger.info("Preview - From: " + senderEmail + ", To: " + recipients + ", Subject: " + subject + ", Date: " + receivedDateString);

                } catch (Exception e) {
                    logger.warning("Error processing message with subject: " + " - " + e.getMessage());
                }

                // Optional: Call garbage collection less frequently
                if (i % 500 == 0) {
                    System.gc();
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing messages in folder: " + sourceFolder.getDisplayName() + " - " + e.getMessage());
        }
    }

    private String getRecipientEmails(MapiMessage message) {
        StringBuilder recipients = new StringBuilder();
        MapiRecipientCollection recipientCollection = message.getRecipients();

        for (int i = 0; i < recipientCollection.size(); i++) {
            MapiRecipient recipient = recipientCollection.get_Item(i);
            MapiProperty smtpAddressProp = recipient.getProperties().get_Item(MapiPropertyTag.PR_SMTP_ADDRESS);

            if (smtpAddressProp != null) {
                recipients.append(smtpAddressProp.getString());
            } else {
                recipients.append(recipient.getEmailAddress());
            }

            if (i < recipientCollection.size() - 1) {
                recipients.append("; ");  // Separate multiple email addresses with a semicolon
            }
        }

        return recipients.toString();
    }


    private void saveAttachmentsToDatabase(Connection conn, MapiMessage message, int messageId) {
        try {
            if (!message.getAttachments().isEmpty()) {
                for (MapiAttachment attachment : message.getAttachments()) {
                    String fileName = attachment.getLongFileName();
                    byte[] fileData = attachment.getBinaryData();
                    // Insert the attachment into the database
                    dbOps.insertAttachment(conn, messageId, fileName, fileData);
                }
            }
        } catch (Exception e) {
            logger.warning("Error saving attachments for message with subject: " + message.getSubject() + " - " + e.getMessage());
        }
    }


    private void processMessage(FolderInfo folder, MapiMessage message) {
        try {
            switch (message.getMessageClass()) {
                case "IPM.Appointment":
                case "IPM.Schedule.Meeting.Request":
                    processCalendar(folder, message);
                    break;
                case "IPM.Contact":
                    processContact(folder, message);
                    break;
                case "IPM.Task":
                    processTask(folder, message);
                    break;
                default:
                    processGenericMessage(folder, message);
                    break;
            }
        } catch (Exception e) {
            logger.warning("Error processing message: " + e.getMessage());
        }
    }

    private void processCalendar(FolderInfo folder, MapiMessage message) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            String input = duplicateCheckCalendar(calendar);
            if (!enableDuplicateCheck || !listdupliccal.contains(input)) {
                synchronized (listdupliccal) {
                    if (enableDuplicateCheck && !listdupliccal.contains(input)) {
                        listdupliccal.add(input);
                    }
                }
                logger.info("Adding calendar to folder: " + folder.getDisplayName());
                folder.addMessage(message);
                incrementCount();
            }
        } catch (Exception e) {
            logger.warning("Error processing calendar: " + e.getMessage());
        }
    }

    private void processContact(FolderInfo folder, MapiMessage message) {
        try {
            MapiContact contact = (MapiContact) message.toMapiMessageItem();
            String input = duplicateCheckContact(contact);
            if (!enableDuplicateCheck || !listdupliccontact.contains(input)) {
                synchronized (listdupliccontact) {
                    if (enableDuplicateCheck && !listdupliccontact.contains(input)) {
                        listdupliccontact.add(input);
                    }
                }
                logger.info("Adding contact to folder: " + folder.getDisplayName());
                folder.addMessage(message);
                incrementCount();
            }
        } catch (Exception e) {
            logger.warning("Error processing contact: " + e.getMessage());
        }
    }

    private void processTask(FolderInfo folder, MapiMessage message) {
        try {
            MapiTask task = (MapiTask) message.toMapiMessageItem();
            String input = duplicateCheckTask(task);
            if (!enableDuplicateCheck || !listduplictask.contains(input)) {
                synchronized (listduplictask) {
                    if (enableDuplicateCheck && !listduplictask.contains(input)) {
                        listduplictask.add(input);
                    }
                }
                logger.info("Adding task to folder: " + folder.getDisplayName());
                folder.addMessage(message);
                incrementCount();
            }
        } catch (Exception e) {
            logger.warning("Error processing task: " + e.getMessage());
        }
    }

    private void processGenericMessage(FolderInfo folder, MapiMessage message) {
        try {
            String input = duplicateCheckMessage(message);
            if (!enableDuplicateCheck || !listduplicacy.contains(input)) {
                synchronized (listduplicacy) {
                    if (enableDuplicateCheck && !listduplicacy.contains(input)) {
                        listduplicacy.add(input);
                    }
                }

                MapiConversionOptions conversionOptions = MapiConversionOptions.getASCIIFormat();
                MailConversionOptions mailConversionOptions = new MailConversionOptions();
                mailConversionOptions.setConvertAsTnef(false);
                MailMessage mailMessage = message.toMailMessage(mailConversionOptions);
                MapiMessage convertedMessage = MapiMessage.fromMailMessage(mailMessage, conversionOptions);
                int bodyType = convertedMessage.getBodyType();
                if (bodyType == 0) {
                    convertedMessage.setBodyContent(convertedMessage.getBodyHtml(), 1);
                } else {
                    convertedMessage.setBodyContent(convertedMessage.getBodyRtf(), 2);
                }

                logger.info("Adding generic message to folder: " + folder.getDisplayName());
                folder.addMessage(convertedMessage);
                incrementCount();
            }
        } catch (Exception e) {
            logger.warning("Error processing generic message: " + e.getMessage());
        }
    }

    private void incrementCount() {
        synchronized (this) {
            count_destination++;
        }
    }

    private String cleanFolderName(String folderName) {
        if (folderName == null || folderName.trim().isEmpty()) {
            return "";
        }
        folderName = folderName.replace(",", "").replace(".", "");
        folderName = folderName.replaceAll("[\\[\\]]", "");
        return folderName.trim();
    }

    private String generateFolderNameFromPath(String path) {
        return path.replace("/", "_").replace("\\", "_").replace(":", "_");
    }

    private String duplicateCheckCalendar(MapiCalendar calendar) {
        String input = calendar.getLocation() + calendar.getStartDate() + calendar.getEndDate();
        return input.replaceAll("\\s", "").trim();
    }

    private String duplicateCheckContact(MapiContact contact) {
        String input = contact.getNameInfo().getDisplayName() + contact.getPersonalInfo().getNotes();
        return input.replaceAll("\\s", "").trim();
    }



    private String duplicateCheckTask(MapiTask task) {
        String input = task.getSubject() + task.getBody();
        return input.replaceAll("\\s", "").trim();
    }

    private String duplicateCheckMessage(MapiMessage message) {
        String input = message.getSubject() + message.getBody();
        return input.replaceAll("\\s", "").trim();
    }
}

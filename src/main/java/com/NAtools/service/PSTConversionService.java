package com.NAtools.service;

import com.NAtools.model.Attachment;
import com.NAtools.model.Folder;
import com.NAtools.model.Message;
import com.aspose.email.*;

import java.sql.Connection;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.logging.Logger;

public class PSTConversionService {

    private static final Logger logger = Logger.getLogger(PSTConversionService.class.getName());
    private final SQLQueryService sqlQueryService = new SQLQueryService();

    public void convertDatabaseToPST(String pstFilePath, Connection connection) {
        try {
            logger.info("Starting PST conversion...");

            PersonalStorage pst = PersonalStorage.create(pstFilePath, FileFormatVersion.Unicode);
            FolderInfo rootFolder = pst.getRootFolder();

            List<Folder> rootFolders = sqlQueryService.getRootFolders(connection);

            for (Folder folder : rootFolders) {
                FolderInfo pstFolder = rootFolder.addSubFolder(folder.getName());
                processSubFoldersAndMessages(pstFolder, folder, connection);
            }

            pst.dispose(); // Ensure PST file is properly closed
            logger.info("PST conversion completed successfully.");
        } catch (Exception e) {
            logger.severe("Error during PST conversion: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processSubFoldersAndMessages(FolderInfo pstFolder, Folder folder, Connection connection) {
        try {
            List<Folder> subFolders = sqlQueryService.getSubFolders(connection, folder.getId());
            for (Folder subFolder : subFolders) {
                FolderInfo newPstFolder = pstFolder.addSubFolder(subFolder.getName());
                processSubFoldersAndMessages(newPstFolder, subFolder, connection);
            }

            List<Message> messages = sqlQueryService.getMessagesForFolder(connection, folder.getId());
            for (Message message : messages) {
                addMessageToPST(pstFolder, message,connection);
            }
        } catch (Exception e) {
            logger.severe("Error processing subfolders and messages: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void addMessageToPST(FolderInfo targetFolder, Message message,Connection connection) {
        try {
            MapiMessage mapiMessage = new MapiMessage();

            // Set subject
            String bodyFormat;
            mapiMessage.setSubject(message.getSubject());

            if ("HTML".equalsIgnoreCase(message.getBodyFormat())) {
                mapiMessage.setBodyContent(message.getBody(), BodyContentType.Html);
            } else if ("RTF".equalsIgnoreCase(message.getBodyFormat())) {
                mapiMessage.setBodyRtf(message.getBody());
            } else {
                mapiMessage.setBody(message.getBody());
            }
            //mapiMessage.setBody(message.getBody());

            // Add recipients to the MapiMessage
            addRecipientsToMapiMessage(mapiMessage, message.getRecipients());

            // Set the sender
            mapiMessage.setSenderEmailAddress(message.getSenderEmail());

            // Set the received date
            try {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                Date date = dateFormat.parse(message.getReceivedDate());
                mapiMessage.setDeliveryTime(date);
            } catch (ParseException e) {
                logger.warning("Error parsing date for message with subject: " + message.getSubject() + ". Using current date.");
                mapiMessage.setDeliveryTime(new Date()); // Default to current date if parsing fails
            }

            // Handle attachments
            addAttachmentsToMapiMessage(mapiMessage, message.getId(),connection);
            // Add the message to the target PST folder
            targetFolder.addMessage(mapiMessage);

            logger.info("Added message with subject: " + message.getSubject() + " to folder: " + targetFolder.getDisplayName());
        } catch (Exception e) {
            logger.severe("Error adding message to PST: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void addRecipientsToMapiMessage(MapiMessage mapiMessage, String recipientEmails) {
        MapiRecipientCollection recipients = new MapiRecipientCollection();
        for (String recipientEmail : recipientEmails.split(";")) {
            if (!recipientEmail.trim().isEmpty()) {
                recipients.add(recipientEmail.trim(), recipientEmail.trim(), MapiRecipientType.MAPI_TO);
            }
        }
        mapiMessage.setRecipients(recipients);
    }
    private void addAttachmentsToMapiMessage(MapiMessage mapiMessage, int messageId,Connection connection) {
        try {
            // Retrieve attachments from the database
            List<com.NAtools.model.Attachment> attachments = sqlQueryService.getAttachmentsForMessage(connection,messageId);
            for (Attachment attachment : attachments) {
                // Assuming attachment.getFilePath() returns the path to the attachment file
                mapiMessage.getAttachments().add(attachment.getName(),attachment.getFileData());
                logger.info("Added attachment: " + attachment.getFileName() + " to message: " + mapiMessage.getSubject());
            }
        } catch (Exception e) {
            logger.warning("Error adding attachments to message with subject: " + mapiMessage.getSubject());
            e.printStackTrace();
        }
    }
}


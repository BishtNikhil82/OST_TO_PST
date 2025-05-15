package com.NAtools.service;

import com.NAtools.Convertor.OSTToPST;
import com.NAtools.model.Message;
import com.aspose.email.FolderInfo;
import com.aspose.email.MapiMessage;
import com.aspose.email.MessageInfo;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import static com.aspose.email.PropertyDataType.Null;

public class UIPreview {

    private static final Logger logger = Logger.getLogger(UIPreview.class.getName());
    private final OSTToPST converter;

    public UIPreview(OSTToPST converter) {
        this.converter = converter;
    }

    // Get the list of folder names for the UI
    public List<String> getFolderList() {
        try {
            return converter.loadInitialFolders();
        } catch (Exception e) {
            logger.severe("Error loading folder list: " + e.getMessage());
            return new ArrayList<>();
        }
    }
    public Map<String, FolderInfo> getFolderMap() {
        try {
            return converter.getLoadedFolders();
        } catch (Exception e) {
            logger.severe("Error loading folder list: " + e.getMessage());
            return null;  // Changed 'Null' to 'null' to return the proper null value in Java
        }
    }


    // Get the list of message previews for a given folder name
    public List<Message> getMessagePreviews(String folderName) {
        List<Message> messagePreviews = new ArrayList<>();
        try {
           // converter.loadFolderContent(folderName);

            FolderInfo folder = converter.getLoadedFolders().get(folderName);
            if (folder != null) {
                folder.getContents().forEach(messageInfo -> {
                    try {
                        MapiMessage message = converter.getSourcePst().extractMessage(messageInfo);
                        if (message != null) {
                            Message preview = new Message();
                            preview.setSubject(message.getSubject());
                            preview.setSenderEmail(message.getSenderEmailAddress());
                            preview.setReceivedDate(message.getDeliveryTime().toString());
                            preview.setRecipients(converter.getRecipientEmails(message));
                            preview.setBody(message.getBody());
                            preview.setBodyFormat(getBodyFormat(message));
                            messagePreviews.add(preview);
                        }
                    } catch (Exception e) {
                        logger.warning("Error extracting message preview: " + e.getMessage());
                    }
                });
                logger.info("Loaded message previews for folder: " + folderName);
            } else {
                logger.warning("Folder not found for preview: " + folderName);
            }

        } catch (Exception e) {
            logger.severe("Error loading message previews for folder: " + folderName + " - " + e.getMessage());
        }

        return messagePreviews;
    }

    // Get the full details of a specific message by subject
    public Message getMessageDetails(String folderName, String subject) {
        try {
            FolderInfo folder = converter.getLoadedFolders().get(folderName);
            if (folder != null) {
                for (MessageInfo messageInfo : folder.getContents()) {
                    try {
                        MapiMessage message = converter.getSourcePst().extractMessage(messageInfo);
                        if (subject.equals(message.getSubject())) {
                            Message detailedMessage = new Message();
                            detailedMessage.setSubject(message.getSubject());
                            detailedMessage.setSenderEmail(message.getSenderEmailAddress());
                            detailedMessage.setReceivedDate(message.getDeliveryTime().toString());
                            detailedMessage.setRecipients(converter.getRecipientEmails(message));
                            detailedMessage.setBody(message.getBody());
                            detailedMessage.setBodyFormat(getBodyFormat(message));
                            return detailedMessage;
                        }
                    } catch (Exception e) {
                        logger.warning("Error extracting message details: " + e.getMessage());
                    }
                }
            } else {
                logger.warning("Folder or message not found: " + folderName + " / " + subject);
            }
        } catch (Exception e) {
            logger.severe("Error retrieving message details: " + e.getMessage());
        }

        return null;
    }

    // Helper method to determine the body format
    private String getBodyFormat(MapiMessage message) {
        try {
            if (message.getBodyHtml() != null) {
                return "HTML";
            } else if (message.getBodyRtf() != null) {
                return "RTF";
            } else {
                return "PlainText";
            }
        } catch (Exception e) {
            logger.warning("Error determining body format: " + e.getMessage());
            return "Unknown";
        }
    }
}

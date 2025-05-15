package com.NAtools.thread;

import com.aspose.email.*;
import java.util.logging.Logger;

public class FolderProcessingTask implements Runnable {
    private static final Logger logger = Logger.getLogger(FolderProcessingTask.class.getName());

    private final FolderInfo sourceFolder;
    private final FolderInfo targetFolder;
    private final PersonalStorage sourcePst;

    public FolderProcessingTask(FolderInfo sourceFolder, FolderInfo targetFolder, PersonalStorage sourcePst) {
        this.sourceFolder = sourceFolder;
        this.targetFolder = targetFolder;
        this.sourcePst = sourcePst;
    }

    @Override
    public void run() {
        processFolder(sourceFolder, targetFolder);
    }

    public void processFolder(FolderInfo sourceFolder, FolderInfo targetFolder) {
        try {
            if (sourceFolder == null || targetFolder == null) {
                logger.warning("Source or target folder is null, skipping folder processing.");
                return;
            }

            logger.info("Processing folder: " + sourceFolder.getDisplayName());

            MessageInfoCollection messageInfoCollection = sourceFolder.getContents();
            for (MessageInfo messageInfo : messageInfoCollection) {
                try {
                    MapiMessage message = sourcePst.extractMessage(messageInfo);
                    processMessage(message, targetFolder);
                } catch (Exception e) {
                    logger.severe("Error processing message: " + e.getMessage());
                    e.printStackTrace();
                }
            }

            for (FolderInfo subFolder : sourceFolder.getSubFolders()) {
                FolderInfo newTargetSubFolder = findOrCreateSubFolder(targetFolder, subFolder.getDisplayName(), subFolder.getContainerClass());
                if (newTargetSubFolder != null) {
                    processFolder(subFolder, newTargetSubFolder);
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing folder: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processMessage(MapiMessage message, FolderInfo targetFolder) {
        if (message == null) {
            logger.warning("Message is null, skipping processing.");
            return;
        }

        try {
            switch (message.getMessageClass()) {
                case "IPM.Note":
                    processEmail(message, targetFolder);
                    break;
                case "IPM.Contact":
                    processContact(message, targetFolder);
                    break;
                case "IPM.Appointment":
                case "IPM.Schedule.Meeting.Request":
                    processCalendar(message, targetFolder);
                    break;
                case "IPM.Task":
                    processTask(message, targetFolder);
                    break;
                default:
                    logger.info("Unsupported message class: " + message.getMessageClass());
            }
        } catch (Exception e) {
            logger.severe("Error processing message class " + message.getMessageClass() + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processEmail(MapiMessage message, FolderInfo targetFolder) {
        try {
            MailConversionOptions options = new MailConversionOptions();
            MailMessage mailMessage = message.toMailMessage(options);
            MapiMessage newMapiMessage = MapiMessage.fromMailMessage(mailMessage, new MapiConversionOptions());
            targetFolder.addMessage(newMapiMessage);
        } catch (Exception e) {
            logger.severe("Error processing email: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processContact(MapiMessage message, FolderInfo targetFolder) {
        try {
            MapiContact contact = (MapiContact) message.toMapiMessageItem();
            targetFolder.addMapiMessageItem(contact);
        } catch (Exception e) {
            logger.severe("Error processing contact: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processCalendar(MapiMessage message, FolderInfo targetFolder) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            targetFolder.addMapiMessageItem(calendar);
        } catch (Exception e) {
            logger.severe("Error processing calendar: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processTask(MapiMessage message, FolderInfo targetFolder) {
        try {
            MapiTask task = (MapiTask) message.toMapiMessageItem();
            targetFolder.addMapiMessageItem(task);
        } catch (Exception e) {
            logger.severe("Error processing task: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private FolderInfo findOrCreateSubFolder(FolderInfo parentFolder, String subFolderName, String containerClass) {
        if (subFolderName == null || subFolderName.isEmpty()) {
            logger.warning("Folder name is null or empty, skipping folder creation.");
            return null;
        }

        for (FolderInfo existingFolder : parentFolder.getSubFolders()) {
            if (existingFolder.getDisplayName().equalsIgnoreCase(subFolderName)) {
                logger.info("Folder already exists: " + subFolderName);
                return existingFolder;
            }
        }

        try {
            logger.info("Creating folder: " + subFolderName);
            FolderInfo newFolder = parentFolder.addSubFolder(subFolderName);
            newFolder.changeContainerClass(containerClass);
            return newFolder;
        } catch (AsposeException e) {
            logger.severe("AsposeException creating subfolder " + subFolderName + ": " + e.getMessage());
            e.printStackTrace();
            return null;
        } catch (com.aspose.email.system.exceptions.ArgumentOutOfRangeException e) {
            logger.severe("ArgumentOutOfRangeException creating subfolder " + subFolderName + ": " + e.getMessage() + " - PageType: " + e.getCause());
            e.printStackTrace();
            return null;
        } catch (Exception e) {
            logger.severe("Error creating subfolder " + subFolderName + ": " + e.getMessage());
            e.printStackTrace();
            return null;
        }
    }
}

package com.NAtools.Convertor;
import com.NAtools.config.LogManagerConfig;
import com.aspose.email.*;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class OSTToPST {
    private static final Logger logger = Logger.getLogger(OSTToPST.class.getName());
    static {
        LogManagerConfig.configureLogger(logger);
    }
    private final PersonalStorage sourcePst;
    private final PersonalStorage targetPst;
    private final String sourcePstPath;
    private final String targetPstPath;
    private final boolean enableDuplicateCheck;
    private final List<String> listduplicacy = new ArrayList<>();
    private final List<String> listdupliccal = new ArrayList<>();
    private final List<String> listduplictask = new ArrayList<>();
    private final List<String> listdupliccontact = new ArrayList<>();
    private final String ostFileName;
    private final String attachmentBasePath;  // Base path for attachments

    // In-memory structure to track loaded folders
    private final Map<String, FolderInfo> loadedFolders = new HashMap<>();

    public OSTToPST(PersonalStorage sourcePst, PersonalStorage targetPst, String sourcePstPath,String targetPstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.targetPst = targetPst;
        this.sourcePstPath = sourcePstPath;
        this.enableDuplicateCheck = enableDuplicateCheck;
        this.ostFileName = new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", "");
        this.attachmentBasePath = sourcePstPath.replaceFirst("[.][^.]+$", "") + "_Attachments";
        this.targetPstPath = targetPstPath;
    }


    // Load initial folders - only the structure, not the content
    public List<String> loadInitialFolders() {
        List<String> folderNames = new ArrayList<>();
        FolderInfo rootFolder = sourcePst.getRootFolder();

        // Recursively load folder names and structure
        loadFoldersRecursively(rootFolder, "", folderNames);

        logger.info("Initial folders loaded: " + folderNames);
        return folderNames;
    }

    private void loadFoldersRecursively(FolderInfo folder, String path, List<String> folderNames) {

        if (folder == null || folder.getSubFolders() == null) {
            logger.warning("Folder or subfolders are null at path: " + path);
            return;
        }
        String folderName = path.isEmpty() ? folder.getDisplayName() : path + "/" + folder.getDisplayName();
        folderNames.add(folderName);
        loadedFolders.put(folderName, folder);

        for (FolderInfo subFolder : folder.getSubFolders()) {
            loadFoldersRecursively(subFolder, folderName, folderNames);
        }
    }
    public Map<String, FolderInfo> getLoadedFolders() {
        return loadedFolders;
    }


    // Load content of a specific folder when the user clicks on it
    public void loadFolderContent(String folderName) {
        FolderInfo folder = loadedFolders.get(folderName);

        if (folder != null) {
            processFolder(folder, targetPst.getRootFolder(),this.sourcePstPath);
        } else {
            logger.warning("Folder not found: " + folderName);
        }
    }

    private void processFolder(FolderInfo sourceFolder, FolderInfo targetFolder, String parentPath) {
        try {
            // Clean and prepare the folder name
            String folderName = cleanFolderName(sourceFolder.getDisplayName());

            // If the folder name is empty, use the parent folder for hierarchy processing
            FolderInfo newTargetFolder = targetFolder;
            String fullPath = parentPath;

            if (!folderName.isEmpty()) {
                // Handle IPM_SUBTREE specific logic
                String newFolderName = folderName;
                if (parentPath.endsWith("IPM_SUBTREE")) {
                    int messageCount = sourceFolder.getContents().size();
                    newFolderName = folderName + " [" + messageCount + "]";
                }

                // Generate the full path for the current folder
                fullPath = parentPath.isEmpty() ? newFolderName : parentPath + "/" + newFolderName;

                // Log the processing information
                logger.info("Processing folder: " + newFolderName + " fullPath: " + fullPath + " contains " + sourceFolder.getContents().size() + " messages.");

                // Create the folder in the target PST
                newTargetFolder = targetFolder.addSubFolder(newFolderName, true);
            } else {
                logger.info("Empty folder name encountered, using parent folder for hierarchy.");
            }

            // Process subfolders recursively
            FolderInfoCollection subFolders = sourceFolder.getSubFolders();
            for (FolderInfo subFolder : subFolders) {
                processFolder(subFolder, newTargetFolder, fullPath);
            }

            // Process messages within the folder
            processMessages(sourceFolder, newTargetFolder,fullPath);

        } catch (Exception e) {
            logger.severe("Error processing folder: " + e.getMessage());
            e.printStackTrace();
        }
    }


    // Process messages in the folder
    private void processMessages(FolderInfo sourceFolder, FolderInfo targetFolder,String fullPath) {
        try {
            MessageInfoCollection messages = sourceFolder.getContents();

            for (int i = 0; i < messages.size(); i++) {
                try {
                    MessageInfo messageInfo = messages.get_Item(i);
                    MapiMessage message = sourcePst.extractMessage(messageInfo);

                    String senderEmail = message.getSenderEmailAddress();
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

                    processMessage(targetFolder, message,fullPath);

                    logger.info("Processed message: Subject - " + message.getSubject());

                } catch (Exception e) {
                    logger.warning("Error processing message: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing messages: " + e.getMessage());
        }
    }

    // Get recipient emails from MapiMessage
    public String getRecipientEmails(MapiMessage message) {
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
                recipients.append("; ");
            }
        }

        return recipients.toString();
    }

    // Process a single message and add it to the PST
    private void processMessage(FolderInfo folder, MapiMessage message,String fullPath) {
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
                    processGenericMessage(folder, message,fullPath);
                    break;
            }
        } catch (Exception e) {
            logger.warning("Error processing message: " + e.getMessage());
        }
    }

    private void processCalendar(FolderInfo folder, MapiMessage message) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            if (!isDuplicateCalendar(calendar)) {
                folder.addMessage(message);
            }
        } catch (Exception e) {
            logger.warning("Error processing calendar: " + e.getMessage());
        }
    }

    private void processContact(FolderInfo folder, MapiMessage message) {
        try {
            MapiContact contact = (MapiContact) message.toMapiMessageItem();
            if (!isDuplicateContact(contact)) {
                folder.addMessage(message);
            }
        } catch (Exception e) {
            logger.warning("Error processing contact: " + e.getMessage());
        }
    }

    private void processTask(FolderInfo folder, MapiMessage message) {
        try {
            MapiTask task = (MapiTask) message.toMapiMessageItem();
            if (!isDuplicateTask(task)) {
                folder.addMessage(message);
            }
        } catch (Exception e) {
            logger.warning("Error processing task: " + e.getMessage());
        }
    }

    private void processGenericMessage(FolderInfo folder, MapiMessage message, String fullPath) {
        try {
            if (!isDuplicateMessage(message)) {
                MailConversionOptions options = new MailConversionOptions();
                options.setConvertAsTnef(false);
                MapiMessage convertedMessage = MapiMessage.fromMailMessage(message.toMailMessage(options));
                folder.addMessage(convertedMessage);
                if (!message.getAttachments().isEmpty()) {
                    // Get the parent directory of the target PST file
                    File targetPstFile = new File(targetPstPath);
                    String baseDirectory = targetPstFile.getParent();

                    // Create a subdirectory under the base directory for the attachments
                    String relativeFolderPath = folder.retrieveFullPath().replaceAll("[:\\\\/*?|<>]", "_"); // Clean folder name to be file-system safe

                    // Construct the full path for the attachment directory
                    File attachmentDir = new File(baseDirectory, "Attachments/" + relativeFolderPath);

                    if (!attachmentDir.exists()) {
                        attachmentDir.mkdirs();
                    }

                    // Save attachments in the newly created directory
                    saveAttachments(message, attachmentDir);
                }

            }
        } catch (Exception e) {
            logger.severe("Error processing generic message attachments: " + e.getMessage());
        }
    }

    private void saveAttachments(MapiMessage message, File attachmentDir) {
        File directory = new File(attachmentDir.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim());
        if (directory.exists()) {
            directory = new File(attachmentDir.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim() + "_" + System.currentTimeMillis());
        }
        directory.mkdir();
        for (MapiAttachment attachment : message.getAttachments()) {
            try {
                attachment.save(directory.getAbsolutePath() + File.separator + cleanFileName(attachment.getDisplayName()).trim());
            } catch (Exception e) {
                attachment.save(directory.getAbsolutePath() + File.separator + cleanFileName(attachment.getLongFileName()).trim());
            }
        }
    }
    // Duplicate check methods (using lists similar to your current implementation)
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
        folderName = folderName.replace(",", "").replace(".", "");
        folderName = folderName.replaceAll("[\\[\\]]", "");
        return folderName.trim();
    }

    private String generateFolderNameFromPath(String path) {
        return path.replace("/", "_").replace("\\", "_").replace(":", "_");
    }

    // Final save operation
    public void savePST() {
        try {
            // Set the display name of the root folder using MAPI properties

            FolderInfo rootFolder = sourcePst.getRootFolder();
            processFolder(rootFolder, targetPst.getRootFolder(),this.sourcePstPath);
            logger.info("Final PST save completed.");
        } catch (Exception e) {
            logger.severe("Error during final PST save: " + e.getMessage());
        } finally {
            sourcePst.dispose();
            targetPst.dispose();
        }
    }
    private String cleanFileName(String fileName) {
        return (fileName != null ? fileName : "Unnamed").replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }
    public PersonalStorage getSourcePst() {
        return this.sourcePst;
    }
}


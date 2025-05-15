package com.NAtools.Convertor;



import com.NAtools.config.LogManagerConfig;
import com.aspose.email.*;
import java.io.File;
import java.util.HashSet;
import java.util.Set;
import java.util.logging.Logger;

public class OSTToEML {
    private static final Logger logger = Logger.getLogger(OSTToEML.class.getName());
    static {
        LogManagerConfig.configureLogger(logger);
    }
    private final PersonalStorage sourcePst;
    private final String outputDirectory;
    private final boolean enableDuplicateCheck;
    private final Set<String> setDuplicacy = new HashSet<>();
    private final Set<String> setDupliccal = new HashSet<>();
    private final Set<String> setDuplictask = new HashSet<>();
    private final Set<String> setDupliccontact = new HashSet<>();
    private final String ostFileName;

    public OSTToEML(PersonalStorage sourcePst, String outputDirectory, String sourcePstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.outputDirectory = outputDirectory;
        this.enableDuplicateCheck = enableDuplicateCheck;

        // Extract the OST file name without the extension
        this.ostFileName = new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", "");
    }

    public void convertToEML() {
        try {
            FolderInfo rootFolder = sourcePst.getRootFolder();
            processFolder(rootFolder, outputDirectory);
            logger.info("Conversion to EML format completed.");
        } catch (Exception e) {
            logger.severe("Error during EML conversion: " + e.getMessage());
        } finally {
            sourcePst.dispose();
        }
    }

    private void processFolder(FolderInfo folder, String parentPath) {
        try {
            String folderName = cleanFolderName(folder.getDisplayName());
            if (folderName.isEmpty()|| folderName.equals("Unnamed")) {
                folderName = ostFileName;
                logger.info("Unnamed folder found. Generated name: " + folderName);
            }

            String fullPath = parentPath.isEmpty() ? folderName : parentPath + "/" + folderName;
            File directory = new File(fullPath);
            if (!directory.exists()) {
                directory.mkdirs();  // Create directories if they don't exist
            }

            FolderInfoCollection subFolders = folder.getSubFolders();
            for (FolderInfo subFolder : subFolders) {
                processFolder(subFolder, fullPath);
            }

            processMessages(folder, fullPath);

        } catch (Exception e) {
            logger.severe("Error processing folder: " + e.getMessage());
        }
    }

    private void processMessages(FolderInfo folder, String folderPath) {
        try {
            MessageInfoCollection messages = folder.getContents();

            for (int i = 0; i < messages.size(); i++) {
                try {
                    MessageInfo messageInfo = messages.get_Item(i);

                    // Create MailConversionOptions object
                    MailConversionOptions options = new MailConversionOptions();
                    options.setConvertAsTnef(false);

                    // Extract the message as a MailMessage
                    MailMessage message = sourcePst.extractMessage(messageInfo).toMailMessage(options);

                    // Determine the type of the message based on content or other properties
                    if (message.getHeaders().contains("X-MS-Has-Attach")) {
                        saveMessage(message, folderPath);
                    } else if (message.getHeaders().contains("X-MS-Calendar")) {
                        saveCalendar(message, folderPath);
                    } else if (message.getHeaders().contains("X-MS-Task")) {
                        saveTask(message, folderPath);
                    } else if (message.getHeaders().contains("X-MS-Contact")) {
                        saveContact(message, folderPath);
                    } else {
                        saveGenericMessage(message, folderPath);
                    }

                } catch (Exception e) {
                    logger.warning("Error processing message: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing messages: " + e.getMessage());
        }
    }


    private void saveMessage(MailMessage message, String folderPath) {
        try {
            if (!isDuplicateMessage(message) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(message.getSubject());
                if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                    System.out.println("Unanamed message with no subject found");
                    baseFileName = "Unnamed";
                }
                String fileName = baseFileName + ".eml";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".eml";
                    file = new File(folderPath, fileName);
                    count++;
                }

                message.save(file.getAbsolutePath(), SaveOptions.getDefaultEml());
                logger.info("Saved email: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving email: " + e.getMessage());
        }
    }

    private void saveContact(MailMessage message, String folderPath) {
        try {
            MapiContact contact = (MapiContact) MapiMessage.fromMailMessage(message).toMapiMessageItem();
            if (!isDuplicateContact(contact)) {
                String fileName = cleanFileName(contact.getNameInfo().getDisplayName()) + ".eml";
                message.save(folderPath + "/" + fileName, SaveOptions.getDefaultEml());
                logger.info("Saved contact: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving contact: " + e.getMessage());
        }
    }

    private void saveCalendar(MailMessage message, String folderPath) {
        try {
            MapiCalendar calendar = (MapiCalendar) MapiMessage.fromMailMessage(message).toMapiMessageItem();
            if (!isDuplicateCalendar(calendar)) {
                String fileName = cleanFileName(calendar.getSubject()) + ".eml";
                message.save(folderPath + "/" + fileName, SaveOptions.getDefaultEml());
                logger.info("Saved calendar: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving calendar: " + e.getMessage());
        }
    }


    private void saveTask(MailMessage message, String folderPath) {
        try {
            MapiTask task = (MapiTask) MapiMessage.fromMailMessage(message).toMapiMessageItem();
            if (!isDuplicateTask(task)) {
                String fileName = cleanFileName(task.getSubject()) + ".eml";
                message.save(folderPath + "/" + fileName, SaveOptions.getDefaultEml());
                logger.info("Saved task: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving task: " + e.getMessage());
        }
    }

    private void saveGenericMessage(MailMessage message, String folderPath) {
        try {
            if (!isDuplicateMessage(message) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(message.getSubject());
                if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                    System.out.println("Unanamed message with no subject found");
                    baseFileName = "Unnamed";
                }
                String fileName = baseFileName + ".eml";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".eml";
                    file = new File(folderPath, fileName);
                    count++;
                }

                message.save(file.getAbsolutePath(), SaveOptions.getDefaultEml());
                logger.info("Saved email: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving email: " + e.getMessage());
        }
    }

    private boolean isDuplicateCalendar(MapiCalendar calendar) {
        String key = calendar.getLocation() + calendar.getStartDate() + calendar.getEndDate();
        if (setDupliccal.contains(key)) {
            return true;
        } else {
            setDupliccal.add(key);
            return false;
        }
    }

    private boolean isDuplicateContact(MapiContact contact) {
        String key = contact.getNameInfo().getDisplayName() + contact.getPersonalInfo().getNotes();
        if (setDupliccontact.contains(key)) {
            return true;
        } else {
            setDupliccontact.add(key);
            return false;
        }
    }

    private boolean isDuplicateTask(MapiTask task) {
        String key = task.getSubject() + task.getBody();
        if (setDuplictask.contains(key)) {
            return true;
        } else {
            setDuplictask.add(key);
            return false;
        }
    }

    private boolean isDuplicateMessage(MailMessage message) {
        String key = message.getSubject() + message.getBody();
        if (setDuplicacy.contains(key)) {
            return true;
        } else {
            setDuplicacy.add(key);
            return false;
        }
    }

    private String cleanFileName(String fileName) {
        return (fileName != null ? fileName : "Unnamed").replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    private String cleanFolderName(String folderName) {
        if (folderName == null || folderName.trim().isEmpty()) {
            return "Unnamed";
        }
        folderName = folderName.replace(",", "").replace(".", "");
        folderName = folderName.replaceAll("[\\[\\]]", "");
        return folderName.trim();
    }
}


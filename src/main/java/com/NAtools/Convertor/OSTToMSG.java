package com.NAtools.Convertor;

import com.NAtools.config.LogManagerConfig;
import com.aspose.email.*;
import java.io.File;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.logging.Logger;

public class OSTToMSG {
    private static final Logger logger = Logger.getLogger(OSTToMSG.class.getName());
    static {
        LogManagerConfig.configureLogger(logger);
    }
    private final PersonalStorage sourcePst;
    private final String outputDirectory;
    private final String sourcePstPath;
    private final boolean enableDuplicateCheck;
    private final Set<String> setDuplicacy = new HashSet<>();
    private final Set<String> setDupliccal = new HashSet<>();
    private final Set<String> setDuplictask = new HashSet<>();
    private final Set<String> setDupliccontact = new HashSet<>();
    private final String ostFileName;

    public OSTToMSG(PersonalStorage sourcePst, String outputDirectory, String sourcePstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.outputDirectory = outputDirectory;
        this.sourcePstPath = sourcePstPath;
        this.enableDuplicateCheck = enableDuplicateCheck;

        // Extract the OST file name without the extension
        this.ostFileName = new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", "");
    }

    public void convertToMSG() {
        try {
            FolderInfo rootFolder = sourcePst.getRootFolder();
            processFolder(rootFolder, outputDirectory);
            logger.info("Conversion to MSG format completed.");
        } catch (Exception e) {
            logger.severe("Error during MSG conversion: " + e.getMessage());
        } finally {
            sourcePst.dispose();
        }
    }

    private void processFolder(FolderInfo folder, String parentPath) {
        try {
            String folderName = cleanFolderName(folder.getDisplayName());
            if (folderName.isEmpty() || folderName.equals("Unnamed")) {
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
                    MapiMessage message = sourcePst.extractMessage(messageInfo);

                    switch (message.getMessageClass()) {
                        case "IPM.Note":
                            saveMessage(message, folderPath);
                            break;
                        case "IPM.Contact":
                            saveContact(message, folderPath);
                            break;
                        case "IPM.Appointment":
                        case "IPM.Schedule.Meeting.Request":
                            saveCalendarAsICS(message, folderPath);
                            break;
                        case "IPM.Task":
                            saveTask(message, folderPath);
                            break;
                        default:
                            saveGenericMessage(message, folderPath);
                            break;
                    }

                } catch (Exception e) {
                    logger.warning("Error processing message: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing messages: " + e.getMessage());
        }
    }

    private void saveMessage(MapiMessage message, String folderPath) {
        try {
            if (!isDuplicateMessage(message) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(message.getSubject());
                if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                    System.out.println("Unanamed message with no subject found");
                    baseFileName = "Unnamed";
                }


                String fileName = baseFileName + ".msg";
                File file = new File(folderPath, fileName);
                int count = 1;

                // If file with the same name already exists, append a number to make it unique
                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".msg";
                    file = new File(folderPath, fileName);
                    count++;
                }

                // Save the message to the uniquely named file
                message.save(file.getAbsolutePath());
                logger.info("Saved email: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving email: " + e.getMessage());
        }
    }


    private void saveContact(MapiMessage message, String folderPath) {
        try {
            MapiContact contact = (MapiContact) message.toMapiMessageItem();
            if (!isDuplicateContact(contact) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(contact.getNameInfo().getDisplayName());
                String fileName = baseFileName + ".msg";
                File file = new File(folderPath, fileName);
                int count = 1;

                // If file with the same name already exists, append a number to make it unique
                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".msg";
                    file = new File(folderPath, fileName);
                    count++;
                }

                // Save the contact to the uniquely named file
                message.save(file.getAbsolutePath());
                logger.info("Saved contact: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving contact: " + e.getMessage());
        }
    }
    private void saveCalendarAsICS(MapiMessage message, String folderPath) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            if (!isDuplicateCalendar(calendar) || !enableDuplicateCheck) {
                // Save as .msg
                String baseFileName = cleanFileName(calendar.getSubject());
                int count = 1;
                // Save as .ics
                String icsFileName = baseFileName + ".ics";
                File icsFile = new File(folderPath, icsFileName);
                count = 1;

                while (icsFile.exists()) {
                    icsFileName = baseFileName + "_" + count + ".ics";
                    icsFile = new File(folderPath, icsFileName);
                    count++;
                }
                // Save the calendar as ICS format
               calendar.save(icsFile.getAbsolutePath(), AppointmentSaveFormat.Ics);
                logger.info("Saved calendar as ICS: " + icsFileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving calendar: " + e.getMessage());
        }
    }



    private void saveCalendarAsMSG(MapiMessage message, String folderPath) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            if (!isDuplicateCalendar(calendar) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(calendar.getSubject());
                String fileName = baseFileName + ".msg";
                File file = new File(folderPath, fileName);
                int count = 1;

                // If file with the same name already exists, append a number to make it unique
                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".msg";
                    file = new File(folderPath, fileName);
                    count++;
                }

                // Save the calendar to the uniquely named file
                message.save(file.getAbsolutePath());
                logger.info("Saved calendar: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving calendar: " + e.getMessage());
        }
    }

    private void saveTask(MapiMessage message, String folderPath) {
        try {
            MapiTask task = (MapiTask) message.toMapiMessageItem();
            if (!isDuplicateTask(task) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(task.getSubject());
                String fileName = baseFileName + ".msg";
                File file = new File(folderPath, fileName);
                int count = 1;

                // If file with the same name already exists, append a number to make it unique
                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".msg";
                    file = new File(folderPath, fileName);
                    count++;
                }

                // Save the task to the uniquely named file
                message.save(file.getAbsolutePath());
                logger.info("Saved task: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving task: " + e.getMessage());
        }
    }


    private void saveGenericMessage(MapiMessage message, String folderPath) {
        try {
            boolean isDuplicate = isDuplicateMessage(message);
            if (!isDuplicate || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(message.getSubject());
                String fileName = baseFileName + ".msg";
                File file = new File(folderPath, fileName);
                int count = 1;

                // Only append a number if the file already exists
                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".msg";
                    file = new File(folderPath, fileName);
                    count++;
                }

                message.save(file.getAbsolutePath());
                logger.info("Saved generic message: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving generic message: " + e.getMessage());
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

    private boolean isDuplicateMessage(MapiMessage message) {
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

    private String generateFolderNameFromPath(String path) {
        return path.replace("/", "_").replace("\\", "_").replace(":", "_");
    }
}

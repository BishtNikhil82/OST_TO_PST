package com.NAtools.service;




import com.aspose.email.*;
import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class PSTtoMboxConverter implements Runnable {
    private long folderMessageCount;
    private List<String> calendarDuplicates;
    private List<String> taskDuplicates;
    private List<String> contactDuplicates;
    private List<String> messageDuplicates;
    private static Date fromDate;
    private static Date toDate;
    private ArrayList<Date> fromDateList;
    private ArrayList<Date> toDateList;
    private String fromFolder;
    private String toFolder;
    private String firstName;
    private String middleName;
    private String lastName;
    private String filePath;
    private String destinationPath;
    private String fileType;
    private long destinationCount;
    private static PersonalStorage pstStorage;
    private String rootFolderName;
    private List<String> pstFolderList;
    private static boolean isDateValid;
    private String tempPath;

    public PSTtoMboxConverter(String fileType, String destinationPath, long destinationCount, String filePath, List<String> pstFolderList, ArrayList<Date> fromDateList, ArrayList<Date> toDateList, String tempPath) {
        this.fileType = fileType;
        this.destinationPath = destinationPath;
        this.destinationCount = destinationCount;
        this.filePath = filePath;
        this.pstFolderList = pstFolderList;
        this.fromDateList = fromDateList;
        this.toDateList = toDateList;
        this.tempPath = tempPath;
        this.calendarDuplicates = new ArrayList<>();
        this.taskDuplicates = new ArrayList<>();
        this.contactDuplicates = new ArrayList<>();
        this.messageDuplicates = new ArrayList<>();
    }

    @Override
    public void run() {
        convertPSTtoMbox();
    }

    private void convertPSTtoMbox() {
        pstStorage = PersonalStorage.fromFile(filePath);
        MailConversionOptions mailConversionOptions = new MailConversionOptions();
        FolderInfo rootFolderInfo = pstStorage.getRootFolder();
        rootFolderName = sanitizeFolderName(rootFolderInfo.getDisplayName());

        initializeFolders(destinationPath, rootFolderName);

        MessageInfoCollection messageInfoCollection = rootFolderInfo.getContents();
        MboxrdStorageWriter mboxWriter = initializeMboxWriter(destinationPath, rootFolderName);

        processMessages(messageInfoCollection, mboxWriter, mailConversionOptions);

        mboxWriter.dispose();

        processSubFolders(rootFolderInfo, rootFolderName);
    }

    private void processMessages(MessageInfoCollection messageInfoCollection, MboxrdStorageWriter mboxWriter, MailConversionOptions mailConversionOptions) {
        int messageSize = messageInfoCollection.size();
        for (int i = 0; i < messageSize; i++) {
            try {
                MessageInfo messageInfo = messageInfoCollection.get_Item(i);
                MapiMessage mapiMessage = pstStorage.extractMessage(messageInfo);
                MailMessage mailMessage = mapiMessage.toMailMessage(mailConversionOptions);

                handleMapiMessage(mboxWriter, mapiMessage, mailMessage);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private void handleMapiMessage(MboxrdStorageWriter mboxWriter, MapiMessage mapiMessage, MailMessage mailMessage) {
        if (mapiMessage.getMessageClass().equals("IPM.Contact")) {
            processContactMessage(mboxWriter, mapiMessage, mailMessage);
        } else if (mapiMessage.getMessageClass().equals("IPM.Appointment") || mapiMessage.getMessageClass().contains("IPM.Schedule.Meeting")) {
            processCalendarMessage(mboxWriter, mapiMessage, mailMessage);
        } else if (mapiMessage.getMessageClass().equals("IPM.Task")) {
            processTaskMessage(mboxWriter, mapiMessage, mailMessage);
        } else {
            writeMessageToMbox(mboxWriter, mailMessage, mapiMessage);
        }
    }

    private void processContactMessage(MboxrdStorageWriter mboxWriter, MapiMessage mapiMessage, MailMessage mailMessage) {
        MapiContact contact = (MapiContact) mapiMessage.toMapiMessageItem();
        setContactDetails(mailMessage, contact);
        String contactKey = generateContactKey(contact);

        if (!contactDuplicates.contains(contactKey)) {
            contactDuplicates.add(contactKey);
            writeMessageToMbox(mboxWriter, mailMessage, mapiMessage);
        }
    }

    private void processCalendarMessage(MboxrdStorageWriter mboxWriter, MapiMessage mapiMessage, MailMessage mailMessage) {
        MapiCalendar calendar = (MapiCalendar) mapiMessage.toMapiMessageItem();
        setCalendarDetails(mailMessage, calendar);
        String calendarKey = generateCalendarKey(calendar);

        if (!calendarDuplicates.contains(calendarKey)) {
            calendarDuplicates.add(calendarKey);
            writeMessageToMbox(mboxWriter, mailMessage, mapiMessage);
        }
    }

    private void processTaskMessage(MboxrdStorageWriter mboxWriter, MapiMessage mapiMessage, MailMessage mailMessage) {
        MapiTask task = (MapiTask) mapiMessage.toMapiMessageItem();
        String taskKey = generateTaskKey(task);

        if (!taskDuplicates.contains(taskKey)) {
            taskDuplicates.add(taskKey);
            writeMessageToMbox(mboxWriter, mailMessage, mapiMessage);
        }
    }

    private void setContactDetails(MailMessage mailMessage, MapiContact contact) {
        mailMessage.setSubject(contact.getSubject() != null ? contact.getSubject() : "");
        mailMessage.setBody(contact.getBody() != null ? contact.getBody() : "");
        mailMessage.setDate(new Date()); // Contacts do not have delivery time

        MapiContactNamePropertySet nameProps = contact.getNameInfo();
        firstName = nameProps.getGivenName() != null ? nameProps.getGivenName() : "";
        middleName = nameProps.getMiddleName() != null ? nameProps.getMiddleName() : "";
        lastName = nameProps.getSurname() != null ? nameProps.getSurname() : "";

        nameProps.setDisplayName(String.join(" ", firstName, middleName, lastName));
        contact.setNameInfo(nameProps);
    }

    private void setCalendarDetails(MailMessage mailMessage, MapiCalendar calendar) {
        mailMessage.setSubject(calendar.getSubject() != null ? calendar.getSubject() : "");
        mailMessage.setHtmlBody(calendar.getBodyHtml() != null ? calendar.getBodyHtml() : "");
        mailMessage.setDate(calendar.getStartDate() != null ? calendar.getStartDate() : new Date());
    }

    private void initializeFolders(String destinationPath, String folderName) {
        String fullPath = destinationPath + File.separator + folderName;
        new File(fullPath).mkdirs();
    }

    private MboxrdStorageWriter initializeMboxWriter(String destinationPath, String folderName) {
        String fullPath = destinationPath + File.separator + folderName + ".mbx";
        return new MboxrdStorageWriter(fullPath, false);
    }

    private void processSubFolders(FolderInfo folderInfo, String parentFolderName) {
        FolderInfoCollection subFolders = folderInfo.getSubFolders();
        for (FolderInfo subFolder : subFolders) {
            processFolder(subFolder, parentFolderName);
        }
    }

    private void processFolder(FolderInfo folderInfo, String parentFolderName) {
        String folderName = sanitizeFolderName(folderInfo.getDisplayName());
        String fullPath = parentFolderName + File.separator + folderName;

        initializeFolders(destinationPath, fullPath);

        MboxrdStorageWriter mboxWriter = initializeMboxWriter(destinationPath, fullPath);

        processMessages(folderInfo.getContents(), mboxWriter, new MailConversionOptions());

        mboxWriter.dispose();

        if (folderInfo.hasSubFolders()) {
            processSubFolders(folderInfo, fullPath);
        }
    }

    private String sanitizeFolderName(String folderName) {
        folderName = folderName.replace(",", "").replace(".", "");
        folderName = folderName.replaceAll("[\\[\\]]", "").trim();
        return folderName.isEmpty() ? "Root Folder" : folderName;
    }

    private String generateContactKey(MapiContact contact) {
        return contact.getNameInfo().getDisplayName().replaceAll("\\s", "").trim();
    }

    private String generateCalendarKey(MapiCalendar calendar) {
        return calendar.getSubject().replaceAll("\\s", "").trim();
    }

    private String generateTaskKey(MapiTask task) {
        return task.getSubject().replaceAll("\\s", "").trim();
    }

    private void writeMessageToMbox(MboxrdStorageWriter mboxWriter, MailMessage mailMessage, MapiMessage mapiMessage) {
        String messageKey = mapiMessage.getSubject().replaceAll("\\s", "").trim();
        if (!messageDuplicates.contains(messageKey)) {
            messageDuplicates.add(messageKey);
            mboxWriter.writeMessage(mailMessage);
            destinationCount++;
        }
    }
}

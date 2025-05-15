package com.NAtools.model;

import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;
import java.util.logging.Logger;

public class LostAndFoundFolder {
    private static final Logger logger = Logger.getLogger(LostAndFoundFolder.class.getName());
    private final PersonalStorage targetPst;
    private FolderInfo lostAndFoundFolder;

    public LostAndFoundFolder(PersonalStorage targetPst) {
        this.targetPst = targetPst;
        initializeLostAndFoundFolder();
    }

    private void initializeLostAndFoundFolder() {
        try {
            lostAndFoundFolder = targetPst.getRootFolder().getSubFolder("Lost and Found");
        } catch (Exception e) {
            lostAndFoundFolder = targetPst.getRootFolder().addSubFolder("Lost and Found");
        }
    }

    public void processFolders(FolderInfo parentFolder, PersonalStorage sourcePst) {
        FolderInfoCollection subFolders = parentFolder.getSubFolders();
        for (int i = 0; i < subFolders.size(); i++) {
            FolderInfo folder = subFolders.get_Item(i);
            if (isLostOrOrphaned(folder)) {
                recoverEmails(folder, sourcePst);
            } else {
                processFolders(folder, sourcePst);
            }
        }
    }

    private boolean isLostOrOrphaned(FolderInfo folder) {
        return folder.getDisplayName().equalsIgnoreCase("Orphaned") || folder.getDisplayName().trim().isEmpty();
    }

    private void recoverEmails(FolderInfo folder, PersonalStorage sourcePst) {
        MessageInfoCollection messages = folder.getContents();
        for (int j = 0; j < messages.size(); j++) {
            MessageInfo messageInfo = messages.get_Item(j);
            MapiMessage message = sourcePst.extractMessage(messageInfo);
            MailConversionOptions mailConversionOptions = new MailConversionOptions();
            MailMessage mailMessage = message.toMailMessage(mailConversionOptions);
            MapiMessage convertedMessage = MapiMessage.fromMailMessage(mailMessage, MapiConversionOptions.getASCIIFormat());
            lostAndFoundFolder.addMessage(convertedMessage);
        }
    }
}

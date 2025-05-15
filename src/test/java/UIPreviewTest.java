import com.NAtools.Convertor.OSTToPST;
import com.NAtools.service.UIPreview;
import com.NAtools.model.Message;
import com.NAtools.util.CleanupUtil;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.PersonalStorage;
import org.junit.*;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

public class UIPreviewTest {

    private static UIPreview uiPreview;
    private static OSTToPST converter;
    private static PersonalStorage targetPst;
    private static PersonalStorage sourcePst;
    private static String pstFilePath;
    private static String sourcePstPath;

    @Before
    public void setUp() {
        pstFilePath = "D:/Gmail_OST_BKUP/pst/output.pst";
        sourcePstPath = "D:/Gmail_OST_BKUP/ost/test5.ost";

        // Clean up the directory before starting the tests
        CleanupUtil.ensureCleanDirectory(new File(pstFilePath).getParentFile());

        // Create the target PST file
        targetPst = PersonalStorage.create(pstFilePath, FileFormatVersion.Unicode);
        targetPst.getStore().changeDisplayName(new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", ""));

        // Load the source OST file
        sourcePst = PersonalStorage.fromFile(sourcePstPath);

        // Initialize the converter and UIPreview classes
        converter = new OSTToPST(sourcePst, targetPst, sourcePstPath, pstFilePath,false);
        uiPreview = new UIPreview(converter);
    }


    public void testGetFolderMap() {
        // Load folders
        converter.loadInitialFolders();
        Map<String, FolderInfo> folders = converter.getLoadedFolders();
        // Assert that the folders are not empty
        assertFalse("Folders map is empty", folders.isEmpty());
    }


    public void testGetFolderList() {
        List<String> folders = uiPreview.getFolderList();
        assertNotNull(folders);
        // Additional assertions can be added based on expected folder names
    }

    public void testGetMessageDetails() {
        // Load folders and get messages from Inbox
        converter.loadInitialFolders();
        List<Message> messages = uiPreview.getMessagePreviews("Root - Mailbox/IPM_SUBTREE/Inbox");
        assertFalse("Inbox folder should not be empty", messages.isEmpty());
    }

    @Test
    public void testGeneratePSTFile() {
        testGetFolderMap();
        testGetFolderList();
        testGetMessageDetails();
        // Generate PST file by saving the data
        converter.savePST();
    }

    @After
    public void tearDown() {
        // Dispose resources after each test
        if (sourcePst != null) {
            sourcePst.dispose();
        }
        if (targetPst != null) {
            targetPst.dispose();
        }

        // Clean up the directory after the tests
       // CleanupUtil.ensureCleanDirectory(new File(pstFilePath).getParentFile());
    }
}

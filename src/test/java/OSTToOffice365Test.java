import com.NAtools.Convertor.OSTToOffice365;
import com.aspose.email.PersonalStorage;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.*;

public class OSTToOffice365Test {

    private OSTToOffice365 converter;
    private PersonalStorage sourcePst;
    private final String sourcePstPath = "D:/Gmail_OST_BKUP/ost/test5.ost";
    private final String username = "tooltest82@outlook.com";
    private final String password = "Samsung@135";

    @Before
    public void setUp() {
        // Initialize the PersonalStorage object for testing
        sourcePst = PersonalStorage.fromFile(sourcePstPath);

        // Initialize the converter with necessary parameters
        converter = new OSTToOffice365(sourcePst, username, password, true);
    }

    @Test
    public void testConnectToOffice365() {
        // Test the connection to Office 365
        assertNotNull("Connection to Office 365 should be established", converter.exchangeClient);
    }

    @Test
    public void testLoadInitialFolders() {
        // Run the method to load initial folders
        converter.loadInitialFolders();

        // Assert that the initial folders have been loaded and created in Office 365
        // Assuming the converter logs or keeps track of loaded folders
        assertFalse("Loaded folders should not be empty", converter.loadedFolders.isEmpty());
    }

    @Test
    public void testProcessFolderAndMessages() {
        // Run the method to process a specific folder and its messages
        converter.loadInitialFolders();

        // Verify that folders and messages have been processed and transferred
        assertTrue("There should be some folders processed", converter.loadedFolders.size() > 0);

        // Optionally, assert specific outcomes, such as checking if specific folders or messages were transferred
        // For example, if you have methods to count messages:
        // assertTrue("There should be some messages transferred", converter.getTransferredMessagesCount() > 0);
    }

    @After
    public void tearDown() {
        // Clean up resources
        if (sourcePst != null) {
            sourcePst.dispose();
        }

        // Optionally, perform any additional cleanup after the test, if needed
    }
}

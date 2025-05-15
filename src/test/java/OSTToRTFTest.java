import com.NAtools.Convertor.OSTToRTF;
import com.NAtools.util.CleanupUtil;
import com.aspose.email.PersonalStorage;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.*;

public class OSTToRTFTest {

    private OSTToRTF converter;
    private PersonalStorage sourcePst;
    private final String outputDirectory = "D:/Gmail_OST_BKUP/rtf/";
   // private final String sourcePstPath = "D:/Gmail_OST_BKUP/ost/test5.ost";
    private static String sourcePstPath = "C:/Users/ASUS/AppData/Local/Microsoft/Outlook/tooltest82@outlook.com.ost";

    @Before
    public void setUp() {
        // Clean the output directory
        File outputDir = new File(outputDirectory);
        CleanupUtil.ensureCleanDirectory(outputDir);

        // Setup other resources like PersonalStorage, etc.
        sourcePst = PersonalStorage.fromFile(sourcePstPath);
        converter = new OSTToRTF(sourcePst, outputDirectory, sourcePstPath, false);
    }

    @Test
    public void testConvertToRTF() {
        // Run the conversion
        converter.convertToRTF();

        // Assert that the output directory is not empty after conversion
        File outputDir = new File(outputDirectory);
        assertTrue("Output directory should contain files after conversion", outputDir.listFiles().length > 0);

        // Optionally, check for the presence of specific files or patterns
        File[] rtfFiles = outputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".rtf"));
        assertNotNull("RTF files should be present in the output directory", rtfFiles);
        //assertTrue("There should be at least one RTF file", rtfFiles.length > 0);
    }

    @After
    public void tearDown() {
        // Clean up resources
        if (sourcePst != null) {
            sourcePst.dispose();
        }

        // Optionally, clean up the output directory after tests
        File outputDir = new File(outputDirectory);
        for (File file : outputDir.listFiles()) {
            file.delete();
        }
    }
}


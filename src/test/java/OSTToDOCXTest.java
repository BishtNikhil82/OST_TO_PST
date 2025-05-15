import com.NAtools.Convertor.OSTToDOCX;  // Import the new OSTToDocx class
import com.NAtools.util.CleanupUtil;
import com.aspose.email.PersonalStorage;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertTrue;

public class OSTToDOCXTest {

    private OSTToDOCX converter; // Use OSTToDocx instead of OSTToMSG
    private PersonalStorage sourcePst;
    private final String outputDirectory = "D:/Gmail_OST_BKUP/docx/"; // Adjusted for DOCX output
    private final String sourcePstPath = "C:/Users/ASUS/AppData/Local/Microsoft/Outlook/tooltest82@outlook.com.ost";

    @Before
    public void setUp() {
        // Clean the output directory
        File outputDir = new File(outputDirectory);
        CleanupUtil.ensureCleanDirectory(outputDir);

        // Setup other resources like PersonalStorage, etc.
        sourcePst = PersonalStorage.fromFile(sourcePstPath);
        converter = new OSTToDOCX(sourcePst, outputDirectory, sourcePstPath, false); // Initialize OSTToDocx
    }

    @Test
    public void testGenerateDocxFiles() {
        converter.convertToDocx(); // Call the method to convert to DOCX

        // Verify that the output directory contains DOCX files
        File outputDir = new File(outputDirectory);
        assertTrue("Output directory should exist and be a directory", outputDir.exists() && outputDir.isDirectory());

        File[] files = outputDir.listFiles();
        assertTrue("There should be at least one DOCX file generated", files != null && files.length > 0);

        boolean docxFilesExist = false;
        for (File file : files) {
            if (file.getName().endsWith(".docx")) {
                docxFilesExist = true;
                break;
            }
        }

        assertTrue("DOCX files should be present in the output directory", docxFilesExist);
    }

    @After
    public void tearDown() {
        // Clean up: close the PersonalStorage
        if (sourcePst != null) {
            sourcePst.dispose();
        }

        // Optional: Remove generated files after the test
        File outputDir = new File(outputDirectory);
        if (outputDir.exists()) {
            for (File file : outputDir.listFiles()) {
                file.delete();
            }
        }
    }
}

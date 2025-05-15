import com.NAtools.Convertor.OSTToMSG;
import com.NAtools.util.CleanupUtil;
import com.aspose.email.PersonalStorage;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertTrue;

public class OSTToMSGTest {

    private OSTToMSG converter;
    private PersonalStorage sourcePst;
    private final String outputDirectory = "D:/Gmail_OST_BKUP/msg/";
    private final String sourcePstPath = "D:/Gmail_OST_BKUP/ost/test5.ost";

    @Before
    public void setUp() {
        // Clean the output directory
        File outputDir = new File(outputDirectory);
        CleanupUtil.ensureCleanDirectory(outputDir);

        // Setup other resources like PersonalStorage, etc.
        sourcePst = PersonalStorage.fromFile(sourcePstPath);
            converter = new OSTToMSG(sourcePst, outputDirectory, sourcePstPath, false);
    }


    @Test
    public void testGenerateMSGFiles() {
        converter.convertToMSG();

        // Verify that the output directory contains MSG files
        File outputDir = new File(outputDirectory);
        assertTrue("Output directory should exist and be a directory", outputDir.exists() && outputDir.isDirectory());

        File[] files = outputDir.listFiles();
        assertTrue("There should be at least one MSG file generated", files != null && files.length > 0);

        boolean msgFilesExist = false;
        for (File file : files) {
            if (file.getName().endsWith(".msg")) {
                msgFilesExist = true;
                break;
            }
        }

       // assertTrue("MSG files should be present in the output directory", msgFilesExist);
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

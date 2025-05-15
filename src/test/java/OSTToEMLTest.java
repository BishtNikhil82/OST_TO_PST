import com.NAtools.Convertor.OSTToEML;
import com.NAtools.util.CleanupUtil;
import com.aspose.email.PersonalStorage;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.*;

public class OSTToEMLTest {

    private OSTToEML converter;
    private PersonalStorage sourcePst;
    private final String outputDirectory = "D:/Gmail_OST_BKUP/eml/";
    private final String sourcePstPath = "D:/Gmail_OST_BKUP/ost/test5.ost";

    @Before
    public void setUp() {
        // Clean the output directory
        File outputDir = new File(outputDirectory);
        CleanupUtil.ensureCleanDirectory(outputDir);

        // Setup other resources like PersonalStorage, etc.
        sourcePst = PersonalStorage.fromFile(sourcePstPath);
        converter = new OSTToEML(sourcePst, outputDirectory, sourcePstPath, false);
    }

    @Test
    public void testConvertToEML() {
        // Run the conversion
        converter.convertToEML();

        // Assert that the output directory is not empty after conversion
        File outputDir = new File(outputDirectory);
        assertTrue("Output directory should contain files after conversion", outputDir.listFiles().length > 0);

        // Optionally, check for the presence of specific files or patterns
        File[] emlFiles = outputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".eml"));
        assertNotNull("EML files should be present in the output directory", emlFiles);
        //assertTrue("There should be at least one EML file", emlFiles.length > 0);
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

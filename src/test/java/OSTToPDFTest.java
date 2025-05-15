import com.NAtools.Convertor.OSTOPDF_Test;
import com.NAtools.Convertor.OSTToPDF;
import com.NAtools.util.CleanupUtil;
import com.aspose.email.FolderInfo;
import com.aspose.email.PersonalStorage;
import org.junit.*;
import java.io.File;
import static org.junit.Assert.*;

public class OSTToPDFTest {

    private static OSTToPDF converter;
    private  static OSTOPDF_Test converter2;
    private static PersonalStorage sourcePst;
    private static String outputDirectory = "D:/Gmail_OST_BKUP/pdf/";
    //private static String sourcePstPath = "D:/Gmail_OST_BKUP/ost/test6.ost";
    private static String sourcePstPath = "C:/Users/ASUS/AppData/Local/Microsoft/Outlook/tooltest82@outlook.com.ost";
    @BeforeClass
    public static void setUpClass() {
        File outputDir = new File(outputDirectory);
        CleanupUtil.ensureCleanDirectory(outputDir);
        // Ensure the output directory is clean before starting the test
        // Load the source OST file
        sourcePst = PersonalStorage.fromFile(sourcePstPath);

        // Initialize the converter class
        //converter = new OSTToPDF(sourcePst, outputDirectory, sourcePstPath, false);
        converter2 = new OSTOPDF_Test(sourcePst, outputDirectory, sourcePstPath, false);
    }

    @Test
    public void testConvertToPDF() {
        // Run the conversion process
        converter2.convertToPDF();

        // After conversion, we should have PDF files in the output directory
        File outputDir = new File(outputDirectory);
        assertTrue("Output directory does not exist!", outputDir.exists());

        // Check if at least one PDF file was created
        File[] pdfFiles = outputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));
        assertTrue("No PDF files were created!", pdfFiles != null && pdfFiles.length > 0);
    }

//    @Test
//    public void testProcessFolder() {
//        FolderInfo rootFolder = sourcePst.getRootFolder();
//        converter.processFolder(rootFolder, outputDirectory);
//
//        // Check if the folder structure is created correctly in the output directory
//        File outputDir = new File(outputDirectory);
//        assertTrue("Root folder was not created in output directory!", outputDir.exists() && outputDir.isDirectory());
//    }
//
//    @Test
//    public void testSaveMessageAsPDF() {
//        FolderInfo rootFolder = sourcePst.getRootFolder();
//        converter.processFolder(rootFolder, outputDirectory);
//
//        // Check if any PDF files are created from the emails
//        File outputDir = new File(outputDirectory);
//        File[] pdfFiles = outputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));
//        assertTrue("No PDF files were saved from messages!", pdfFiles != null && pdfFiles.length > 0);
//    }

    @AfterClass
    public static void tearDownClass() {
        // Clean up resources if necessary
        if (sourcePst != null) {
            sourcePst.dispose();
        }
    }
}

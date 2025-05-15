package com.NAtools;

import com.NAtools.Convertor.OSTToPST;
import com.NAtools.db.DBConnection;
import com.NAtools.service.EmailQueryService;
import com.NAtools.service.FolderHierarchyService;

import com.NAtools.service.PSTConversionService;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.PersonalStorage;

import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class Main {
    private static final Logger logger = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) throws Exception {
//        String sourcePstPath = "D:/Gmail_OST_BKUP/ost/test5.ost";
//        String targetPstPath = "D:/Gmail_OST_BKUP/pst/test_mypath.pst";
//
//        PersonalStorage sourcePst = null;
//        PersonalStorage targetPst = null;
//        try {
//            sourcePst = PersonalStorage.fromFile(sourcePstPath);
//            targetPst = PersonalStorage.create(targetPstPath, FileFormatVersion.Unicode);
//
//            // Create the FolderHierarchyService instance
//            OSTToPST convertor_OSTToPST = new OSTToPST(sourcePst, targetPst, sourcePstPath, false);
//            convertor_OSTToPST.loadInitialFolders();
//
//            Map<String, FolderInfo> mm = convertor_OSTToPST.getLoadedFolders();
//
//            System.err.println("Intial Folder List check : ");
//
//        } catch (Exception e) {
//            System.err.println("An error occurred during PST processing: " + e.getMessage());
//            e.printStackTrace();
//        } finally {
//            if (sourcePst != null) {
//                sourcePst.dispose();
//            }
//            if (targetPst != null) {
//                targetPst.dispose();
//            }
//        }

//        try {
//            sourcePst = PersonalStorage.fromFile(sourcePstPath);
//            targetPst = PersonalStorage.create(targetPstPath, FileFormatVersion.Unicode);
//
//            // Create the FolderHierarchyService instance
//            FolderHierarchyService service = new FolderHierarchyService(sourcePst, targetPst, sourcePstPath, false);
//
//            // Run the service in a new thread
//            Thread thread = new Thread(service);
//            thread.start();
//
//            // Optional: Wait for the thread to finish if needed
//            thread.join();
//
//        } catch (Exception e) {
//            System.err.println("An error occurred during PST processing: " + e.getMessage());
//            e.printStackTrace();
//        } finally {
//            if (sourcePst != null) {
//                sourcePst.dispose();
//            }
//            if (targetPst != null) {
//                targetPst.dispose();
//            }
//        }
//        try (Connection connection = DBConnection.getInstance().getConnection()) {
//            String TTPst = "D:/Gmail_OST_BKUP/pst/output1.pst";
//            PSTConversionService conversionService = new PSTConversionService();
//            conversionService.convertDatabaseToPST(TTPst,connection);
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }



//
//        // Queries test on DB
//        Connection conn = DBConnection.getInstance().getConnection();
//        EmailQueryService emailQueryService = new EmailQueryService(conn);
//
//        // Example: Get emails by date range and folder
//        String startDate = "2024-06-01 00:00:00";  // ISO string date
//        String endDate = "2024-08-10 23:59:59";    // ISO string date
//
//        List<String> emailsByDateAndFolder = emailQueryService.getEmailsByDateRangeAndFolder("Inbox", startDate, endDate);
//
//        for (String subject : emailsByDateAndFolder) {
//            logger.severe("Email Subject: " + subject);
//        }
//
//        // Example: Get emails by sender
//        List<String> emailsBySender = emailQueryService.getEmailsBySender("example@example.com");
//        for (String subject : emailsBySender) {
//            logger.info("Email Subject: " + subject);
//        }
//
//        // Example: Get emails by subject keyword
//        List<String> emailsByKeyword = emailQueryService.getEmailsBySubjectKeyword("Project");
//        for (String subject : emailsByKeyword) {
//            logger.info("Email Subject: " + subject);
//        }
    }
}

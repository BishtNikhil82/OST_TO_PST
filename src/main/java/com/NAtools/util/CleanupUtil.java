package com.NAtools.util;
import java.io.File;

public class CleanupUtil {

    public static void cleanDirectory(File directory) {
        if (directory.exists()) {
            File[] files = directory.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isDirectory()) {
                        cleanDirectory(file); // Recursively clean subdirectories
                    }
                    file.delete(); // Delete files and empty directories
                }
            }
        }
    }

    public static void ensureCleanDirectory(File directory) {
        cleanDirectory(directory);
        if (!directory.exists()) {
            directory.mkdirs();  // Ensure the directory exists after cleaning
        }
    }
}

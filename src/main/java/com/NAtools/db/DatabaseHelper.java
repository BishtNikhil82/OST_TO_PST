package com.NAtools.db;
import com.NAtools.config.ConfigLoader;
import com.NAtools.config.*;
import com.NAtools.service.FolderHierarchyService;

import java.io.*;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Logger;

public class DatabaseHelper {
    private static final Logger logger = Logger.getLogger(DatabaseHelper.class.getName());
    static {
        LogManagerConfig.configureLogger(logger);
    }
    static ConfigLoader configLoader = ConfigLoader.getInstance("config.properties");
    private static final String sqlScript = ConfigLoader.getProperty("sql.script");

    private static void createNewDatabase() {
        try {
              Connection conn = DBConnection.getInstance().getConnection();
            if (conn != null) {
                logger.info("A new database has been created at");
                conn.close();
            }
        } catch (Exception e) {
            logger.severe(e.getMessage());
        }
    }

    public static void executeSqlScript() {
        createNewDatabase();
        Connection conn = null;
        try {
            conn = DBConnection.getInstance().getConnection();
            try (Statement stmt = conn.createStatement();
                 InputStream is = DatabaseHelper.class.getClassLoader().getResourceAsStream(sqlScript);
                 BufferedReader br = new BufferedReader(new InputStreamReader(is))) {

                String line;
                StringBuilder sql = new StringBuilder();
                while ((line = br.readLine()) != null) {
                    sql.append(line).append("\n");
                    if (line.trim().endsWith(";")) {  // Ensure that statements are terminated with semicolons
                        logger.info("Executing SQL: " + sql.toString());
                        stmt.executeUpdate(sql.toString());
                        sql.setLength(0); // Reset the StringBuilder for the next statement
                    }
                }
                logger.info("SQL script executed successfully.");
            } catch (SQLException | IOException e) {
                logger.severe("Error executing SQL script: " + e.getMessage());
                e.printStackTrace();
            }
        } finally {
            if (conn != null) {
                try {
                    conn.close();  // Close the connection
                    logger.info("Database connection closed after executing SQL script.");
                } catch (SQLException e) {
                    logger.severe("Error closing database connection: " + e.getMessage());
                    e.printStackTrace();
                }
            }
        }
    }



}

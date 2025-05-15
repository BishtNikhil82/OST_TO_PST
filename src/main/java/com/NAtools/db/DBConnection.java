package com.NAtools.db;

import com.NAtools.config.ConfigLoader;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DBConnection {
    private static DBConnection instance;
    private Connection connection;

    private DBConnection() {
        try {
            ConfigLoader configLoader = ConfigLoader.getInstance("config.properties");
            String dbUrl = configLoader.getProperty("db.url");

            // Ensure the directory exists
            File dbFile = new File(dbUrl);
            File directory = dbFile.getParentFile();
            if (directory != null && !directory.exists()) {
                directory.mkdirs();
            }

        } catch (Exception e) {
            System.out.println("Failed to create the database connection.");
            e.printStackTrace();
        }
    }

    public static synchronized DBConnection getInstance() {
        if (instance == null) {
            instance = new DBConnection();
        }
        return instance;
    }

    public Connection getConnection() {
        try {
            if (connection == null || connection.isClosed()) {
                System.out.println("******* CONNECTING  TO DATABASE.************");
                ConfigLoader configLoader = ConfigLoader.getInstance("config.properties");
                String dbUrl = configLoader.getProperty("db.url");
                connection = DriverManager.getConnection(dbUrl);
            }
        } catch (SQLException e) {
            System.out.println("Failed to get the database connection.");
            e.printStackTrace();
        }
        return connection;
    }
}

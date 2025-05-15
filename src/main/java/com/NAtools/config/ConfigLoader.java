package com.NAtools.config;


import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ConfigLoader {
    private static final Logger logger = Logger.getLogger(ConfigLoader.class.getName());

    static {
        LogManagerConfig.configureLogger(logger);
    }

    private static ConfigLoader instance;
    private static Properties properties;

    private ConfigLoader(String propertiesFileName) {
        properties = new Properties();
        try (InputStream input = getClass().getClassLoader().getResourceAsStream(propertiesFileName)) {
            if (input == null) {
                logger.severe("Sorry, unable to find " + propertiesFileName);
                return;
            }
            properties.load(input);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    public static ConfigLoader getInstance(String propertiesFileName) {
        if (instance == null) {
            synchronized (ConfigLoader.class) {
                if (instance == null) {
                    instance = new ConfigLoader(propertiesFileName);
                }
            }
        }
        return instance;
    }

    public static Level getLogLevel() {
        String logLevel = properties.getProperty("log.level", "INFO");
        return Level.parse(logLevel.toUpperCase());

    }
    public static String getLogPath() {
        String logDirectory = properties.getProperty("log.path", "logs");

        // Ensure the directory exists
        try {
            Files.createDirectories(Paths.get(logDirectory));
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Generate the log file name with the current date
        String date = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        return Paths.get(logDirectory, "application_" + date + ".log").toString();
    }
    public static String getProperty(String key) {
        return properties.getProperty(key);
    }
}

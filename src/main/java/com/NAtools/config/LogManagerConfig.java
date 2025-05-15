package com.NAtools.config;

import java.io.IOException;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.Level;
import java.util.logging.SimpleFormatter;

public class LogManagerConfig {
    private static FileHandler fileHandler;

    public static void configureLogger(Logger logger) {
        ConfigLoader cf = ConfigLoader.getInstance("config.properties");
        Level level = ConfigLoader.getLogLevel();
        logger.setLevel(level);

        if (fileHandler == null) {
            try {
                String logPath = ConfigLoader.getLogPath();
                System.out.println("Log path: " + logPath);
                fileHandler = new FileHandler(logPath, true) {
                    @Override
                    public synchronized void publish(java.util.logging.LogRecord record) {
                        super.publish(record);
                        flush();  // Flush after every log entry
                    }
                };
                fileHandler.setLevel(level);
                fileHandler.setFormatter(new SimpleFormatter());
                logger.addHandler(fileHandler);

                // Attach to root logger
                Logger rootLogger = Logger.getLogger("");
                rootLogger.addHandler(fileHandler);
                rootLogger.setLevel(level);

                System.out.println("FileHandler added to logger.");
                logger.info("This is a test log to verify file logging.");

                // Add a shutdown hook to ensure logs are flushed and closed on shutdown
                Runtime.getRuntime().addShutdownHook(new Thread(() -> {
                    if (fileHandler != null) {
                        fileHandler.flush();
                        fileHandler.close();
                    }
                }));

            } catch (IOException e) {
                logger.severe("Failed to set up file logging: " + e.getMessage());
                e.printStackTrace();
            }
        }

        // Debugging: List all handlers attached to the logger
        for (java.util.logging.Handler handler : logger.getHandlers()) {
            System.out.println("Handler attached: " + handler.getClass().getName());
        }
    }
}
